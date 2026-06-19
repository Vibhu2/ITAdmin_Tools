function Get-VBActiveDirectoryReport {
    <#
    .SYNOPSIS
        Generates Active Directory report
    
    .DESCRIPTION
        Collects and displays Active Directory environment information
    
    .PARAMETER ComputerName
        Target computer name (defaults to local computer)
    
    .EXAMPLE
        Get-ActiveDirectoryReport
        
    .EXAMPLE
        Get-ActiveDirectoryReport -ComputerName "DC01"
    #>
    
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$ComputerName = $env:COMPUTERNAME
    )
    
    # Helper function to get Azure AD join status
    function Get-AzureADJoinStatus {
        param ([string]$TargetComputer = $env:COMPUTERNAME)
        
        try {
            $scriptBlock = {
                try {
                    $dsregResult = dsregcmd /status 2>&1
                    if ($dsregResult -match "AzureAdJoined\s*:\s*YES") {
                        return "Yes"
                    }
                    elseif ($dsregResult -match "DomainJoined\s*:\s*YES") {
                        return "Domain Joined (Hybrid possible)"
                    }
                    else {
                        return "No"
                    }
                }
                catch {
                    return "Unable to determine"
                }
            }
            
            if ($TargetComputer -eq $env:COMPUTERNAME) {
                return & $scriptBlock
            }
            else {
                return Invoke-Command -ComputerName $TargetComputer -ScriptBlock $scriptBlock -ErrorAction SilentlyContinue
            }
        }
        catch {
            return "Error checking status"
        }
    }
    
    try {
        Write-Host "Active Directory Report - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Yellow
        Write-Host ""
        
        # Check and import Active Directory module
        if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
            Write-Host "ERROR: Active Directory module not available" -ForegroundColor Red
            return
        }
        
        if (-not (Get-Module -Name ActiveDirectory)) {
            Import-Module ActiveDirectory -ErrorAction Stop
        }
        
        # Azure AD Join Status
        Write-Host "Azure AD Join Status" -ForegroundColor Cyan
        try {
            $azureADStatus = Get-AzureADJoinStatus -TargetComputer $ComputerName
            $azureTable = @(
                [PSCustomObject]@{
                    'Azure AD Joined' = $azureADStatus
                }
            )
            $azureTable | Format-Table -AutoSize
        }
        catch {
            Write-Host "ERROR: Unable to check Azure AD join status - $($_.Exception.Message)" -ForegroundColor Red
        }
        
        # FSMO Roles Information
        Write-Host "FSMO Roles Information" -ForegroundColor Cyan
        try {
            $domain = Get-ADDomain
            $forest = Get-ADForest
            
            $fsmoTable = @(
                [PSCustomObject]@{ 'FSMO Role' = 'Schema Master'; 'Current Holder' = $forest.SchemaMaster }
                [PSCustomObject]@{ 'FSMO Role' = 'Domain Naming Master'; 'Current Holder' = $forest.DomainNamingMaster }
                [PSCustomObject]@{ 'FSMO Role' = 'PDC Emulator'; 'Current Holder' = $domain.PDCEmulator }
                [PSCustomObject]@{ 'FSMO Role' = 'RID Master'; 'Current Holder' = $domain.RIDMaster }
                [PSCustomObject]@{ 'FSMO Role' = 'Infrastructure Master'; 'Current Holder' = $domain.InfrastructureMaster }
            )
            
            $fsmoTable | Format-Table -AutoSize
        }
        catch {
            Write-Host "ERROR: Unable to retrieve FSMO roles - $($_.Exception.Message)" -ForegroundColor Red
        }
        
        # Functional Levels
        Write-Host "Functional Levels" -ForegroundColor Cyan
        try {
            $domainFunctionalLevel = (Get-ADDomain).DomainMode
            $forestFunctionalLevel = (Get-ADForest).ForestMode
            
            $functionalLevelsTable = @(
                [PSCustomObject]@{
                    'Domain Functional Level' = $domainFunctionalLevel
                    'Forest Functional Level' = $forestFunctionalLevel
                }
            )
            $functionalLevelsTable | Format-Table -AutoSize
        }
        catch {
            Write-Host "ERROR: Unable to retrieve functional levels - $($_.Exception.Message)" -ForegroundColor Red
        }
        
        # Active Directory Recycle Bin Status
        Write-Host "Active Directory Recycle Bin Status" -ForegroundColor Cyan
        try {
            $recycleBinEnabled = (Get-ADOptionalFeature -Filter 'Name -eq "Recycle Bin Feature"').EnabledScopes.Count -gt 0
            $recycleBinStatus = if ($recycleBinEnabled) { "Enabled" } else { "Disabled" }
            
            $recycleBinTable = @(
                [PSCustomObject]@{
                    'AD Recycle Bin Enabled' = $recycleBinStatus
                }
            )
            $recycleBinTable | Format-Table -AutoSize
        }
        catch {
            Write-Host "ERROR: Unable to check recycle bin status - $($_.Exception.Message)" -ForegroundColor Red
        }
        
        # Tombstone Lifetime
        Write-Host "Tombstone Lifetime" -ForegroundColor Cyan
        try {
            $configDN = (Get-ADRootDSE).configurationNamingContext
            $tombstoneLifetime = (Get-ADObject -Identity "CN=Directory Service,CN=Windows NT,CN=Services,$configDN" -Properties tombstoneLifetime).tombstoneLifetime
            
            $tombstoneTable = @(
                [PSCustomObject]@{
                    'Tombstone Lifetime (days)' = $tombstoneLifetime
                }
            )
            $tombstoneTable | Format-Table -AutoSize
        }
        catch {
            Write-Host "ERROR: Unable to retrieve tombstone lifetime - $($_.Exception.Message)" -ForegroundColor Red
        }
        
        # Domain Controllers
        Write-Host "Domain Controllers" -ForegroundColor Cyan
        try {
            $domainControllers = Get-ADDomainController -Filter * | Select-Object Name, Site, IPv4Address, OperatingSystem
            
            $dcTable = foreach ($dc in $domainControllers) {
                $status = if (Test-Connection -ComputerName $dc.Name -Count 1 -Quiet -ErrorAction SilentlyContinue) { "Online" } else { "Offline" }
                
                [PSCustomObject]@{
                    'Domain Controller Name' = $dc.Name
                    'Status' = $status
                    'Site' = $dc.Site
                    'IP Address' = $dc.IPv4Address
                    'Operating System' = $dc.OperatingSystem
                }
            }
            
            $dcTable | Format-Table -AutoSize
        }
        catch {
            Write-Host "ERROR: Unable to retrieve domain controllers - $($_.Exception.Message)" -ForegroundColor Red
        }
        
        # All Servers
        Write-Host "All Servers" -ForegroundColor Cyan
        try {
            $allServers = Get-ADComputer -Filter { OperatingSystem -Like "Windows Server*" } -Property Name, IPv4Address, OperatingSystem, Enabled, LastLogonDate | 
                         Sort-Object Name
            
            $serverTable = foreach ($server in $allServers) {
                $status = if ($server.Enabled) {
                    if ($server.LastLogonDate -and $server.LastLogonDate -gt (Get-Date).AddDays(-30)) {
                        "Online"
                    } else {
                        "Enabled (Inactive)"
                    }
                } else {
                    "Disabled"
                }
                
                [PSCustomObject]@{
                    'Server Name' = $server.Name
                    'Status' = $status
                    'IP Address' = $server.IPv4Address
                    'Operating System' = $server.OperatingSystem
                    'Last Logon' = if ($server.LastLogonDate) { $server.LastLogonDate.ToString("yyyy-MM-dd") } else { "Never" }
                }
            }
            
            $serverTable | Format-Table -AutoSize
        }
        catch {
            Write-Host "ERROR: Unable to retrieve server information - $($_.Exception.Message)" -ForegroundColor Red
        }
        
        # SYSVOL Scripts
        Write-Host "SYSVOL Scripts" -ForegroundColor Cyan
        try {
            $scriptBlock = {
                if (Test-Path -Path "C:\Windows\SYSVOL\sysvol") {
                    $folderpath = (Get-ChildItem "C:\Windows\SYSVOL\sysvol" | Where-Object { $_.PSIsContainer } | Select-Object -First 1).FullName
                    Get-ChildItem -Recurse -Path "$folderpath" -ErrorAction SilentlyContinue | 
                        Where-Object { $_.Extension -in ".bat", ".cmd", ".ps1", ".vbs", ".exe", ".msi" } | 
                        Select-Object FullName, @{Name='Size(KB)';Expression={[math]::Round($_.Length/1KB,2)}}, LastWriteTime, Extension |
                        Sort-Object Extension, Name
                } else {
                    return $null
                }
            }
            
            if ($ComputerName -eq $env:COMPUTERNAME) {
                $sysvolScripts = & $scriptBlock
            } else {
                $sysvolScripts = Invoke-Command -ComputerName $ComputerName -ScriptBlock $scriptBlock -ErrorAction SilentlyContinue
            }
            
            if ($sysvolScripts -and $sysvolScripts.Count -gt 0) {
                $sysvolTable = foreach ($script in $sysvolScripts) {
                    [PSCustomObject]@{
                        'Script Name' = Split-Path $script.FullName -Leaf
                        'Full Path' = $script.FullName
                        'Size (KB)' = $script.'Size(KB)'
                        'Extension' = $script.Extension
                        'Last Modified' = $script.LastWriteTime.ToString("yyyy-MM-dd HH:mm")
                    }
                }
                
                $sysvolTable | Format-Table -AutoSize
            } else {
                Write-Host "No scripts found in SYSVOL or SYSVOL path not accessible." -ForegroundColor Yellow
            }
        }
        catch {
            Write-Host "ERROR: Unable to retrieve SYSVOL scripts - $($_.Exception.Message)" -ForegroundColor Red
        }
        
        # Total AD Users
        Write-Host "Total AD Users" -ForegroundColor Cyan
        try {
            $totalUsers = (Get-ADUser -Filter *).Count
            $enabledUsers = (Get-ADUser -Filter { Enabled -eq $true }).Count
            $disabledUsers = (Get-ADUser -Filter { Enabled -eq $false }).Count
            
            $userStatsTable = @(
                [PSCustomObject]@{
                    'Total AD Users' = $totalUsers
                    'Enabled Users' = $enabledUsers
                    'Disabled Users' = $disabledUsers
                }
            )
            $userStatsTable | Format-Table -AutoSize
        }
        catch {
            Write-Host "ERROR: Unable to retrieve user statistics - $($_.Exception.Message)" -ForegroundColor Red
        }
        
        Write-Host "Report completed successfully!" -ForegroundColor Green
        
    }
    catch {
        Write-Host "CRITICAL ERROR: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Example usage:
 Get-VBActiveDirectoryReport
# Get-ActiveDirectoryReport -ComputerName "DC01"
