# ============================================================
# FUNCTION : Dismount-VBUserHive
# MODULE   : VB.WorkstationReport
# VERSION  : 1.0.1
# CHANGED  : 23-04-2026 -- Finalized: expanded help block with remote, WhatIf, and finally-block examples
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Safely unloads a user registry hive mounted by Mount-VBUserHive
# ENCODING : UTF-8 with BOM
# ============================================================

function Dismount-VBUserHive {
    <#
    .SYNOPSIS
        Safely unloads a user registry hive previously mounted by Mount-VBUserHive.

    .DESCRIPTION
        Dismount-VBUserHive unloads a user NTUSER.DAT hive from HKEY_USERS.
        It only unloads hives that were actively mounted in the current operation
        (HiveMounted = $true). Hives that were already loaded before Mount-VBUserHive
        was called (AlreadyLoaded = $true) are left untouched.

        Designed to accept pipeline input directly from Mount-VBUserHive output.
        Forces a garbage collection pass before unloading to release any open handles
        that would cause reg.exe to fail with "Access is denied".

        Supports local and remote execution via ComputerName.

    .PARAMETER SID
        The SID of the user hive to unload. Maps to HKU\{SID}.

    .PARAMETER HiveMounted
        When true, the unload is performed. When false (hive was already loaded
        before mounting), the function skips the unload and returns Skipped status.
        Accepts pipeline input from Mount-VBUserHive.

    .PARAMETER ComputerName
        Target computer. Defaults to local machine.

    .PARAMETER Credential
        Credentials for remote execution. Not required for local or domain-joined targets.

    .EXAMPLE
        Mount-VBUserHive -Username 'jdoe' | Dismount-VBUserHive

        Mounts and then safely dismounts the hive for jdoe. Skips the unload and
        returns Status = 'Skipped' if the hive was already loaded before the mount call
        (i.e. the user was already logged on).

    .EXAMPLE
        $mount = Mount-VBUserHive -SID 'S-1-5-21-...'
        try {
            # ... registry work against HKU\{SID} ...
        }
        finally {
            $mount | Dismount-VBUserHive | Out-Null
        }

        Recommended pattern -- use a finally block to guarantee dismount even if
        registry work throws an error.

    .EXAMPLE
        Mount-VBUserHive -Username 'jdoe' | Dismount-VBUserHive -WhatIf

        Dry run -- shows what would be unloaded without performing the unload.
        Useful for validating which hives are mounted before making changes.

    .EXAMPLE
        $cred  = Get-Credential
        $mount = Mount-VBUserHive -Username 'jdoe' -ComputerName 'WS001' -Credential $cred
        # ... remote registry work ...
        $mount | Dismount-VBUserHive -Credential $cred

        Remote execution -- mounts and dismounts a hive on a remote workstation.
        Credential is forwarded to the Invoke-Command call inside Dismount-VBUserHive.

    .OUTPUTS
        PSCustomObject
        Returns one object per call with:
          - ComputerName  : Target computer
          - SID           : User SID that was processed
          - HiveUnloaded  : True if unload was performed
          - Skipped       : True if unload was skipped (AlreadyLoaded or HiveMounted false)
          - Status        : 'Success', 'Skipped', or 'Failed'
          - Error         : Error message (only present on failure)

    .NOTES
        Version  : 1.0.1
        Author   : Vibhu Bhatnagar
        Category : User Profile Management

        Requirements:
        - PowerShell 5.1 or higher
        - Administrative privileges (reg.exe load/unload requires elevation)
        - No open handles to the hive being unloaded
    #>

    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$SID,

        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [bool]$HiveMounted = $true,

        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$ComputerName = $env:COMPUTERNAME,

        [PSCredential]$Credential
    )

    process {
        # Skip -- hive was already loaded before we mounted it, leave it alone
        if (-not $HiveMounted) {
            return [PSCustomObject]@{
                ComputerName = $ComputerName
                SID          = $SID
                HiveUnloaded = $false
                Skipped      = $true
                Status       = 'Skipped'
            }
        }

        if (-not $PSCmdlet.ShouldProcess("HKU\$SID on $ComputerName", 'Unload registry hive')) {
            return
        }

        $scriptBlock = {
            param($SID)

            # Flush handles before unload -- prevents "Access is denied" from reg.exe
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            Start-Sleep -Seconds 1

            # Confirm hive is still loaded before attempting unload
            $loadedSIDs = Get-ChildItem 'Registry::HKEY_USERS' |
                Select-Object -ExpandProperty PSChildName

            if ($loadedSIDs -notcontains $SID) {
                return @{ AlreadyGone = $true }
            }

            $regResult = reg.exe unload "HKU\$SID" 2>&1

            if ($LASTEXITCODE -ne 0) {
                throw "reg.exe unload failed: $regResult"
            }

            return @{ AlreadyGone = $false }
        }

        try {
            $result = if ($ComputerName -eq $env:COMPUTERNAME) {
                & $scriptBlock $SID
            }
            else {
                $invokeParams = @{
                    ComputerName = $ComputerName
                    ScriptBlock  = $scriptBlock
                    ArgumentList = $SID
                    ErrorAction  = 'Stop'
                }
                if ($Credential) { $invokeParams['Credential'] = $Credential }
                Invoke-Command @invokeParams
            }

            [PSCustomObject]@{
                ComputerName = $ComputerName
                SID          = $SID
                HiveUnloaded = -not $result.AlreadyGone
                Skipped      = $result.AlreadyGone
                Status       = 'Success'
            }
        }
        catch {
            [PSCustomObject]@{
                ComputerName = $ComputerName
                SID          = $SID
                HiveUnloaded = $false
                Skipped      = $false
                Error        = $_.Exception.Message
                Status       = 'Failed'
            }
        }
    }
}
