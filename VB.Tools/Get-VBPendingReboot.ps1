function Get-VBPendingReboot {

    <#

    .SYNOPSIS

    Checks various registry locations and system flags to determine if Windows is pending a reboot.



    .DESCRIPTION

    This function examines multiple Windows subsystems to detect pending reboot conditions including:

    - Component-Based Servicing (CBS)

    - Windows Update

    - Pending file rename operations

    - Server Feature Manager (Server roles/features)

    - Computer rename operations

    - Domain join operations

    - Optional: SCCM client reboot flags



    .PARAMETER ComputerName

    Specifies the computer name(s) to check. Defaults to local computer.



    .PARAMETER IncludeSCCM

    Include SCCM/ConfigMgr client reboot detection (requires SCCM client).



    .PARAMETER Credential

    Specifies credentials for remote computer access.



    .EXAMPLE

    Get-VBPendingReboot

    Checks the local computer for pending reboots.



    .EXAMPLE

    Get-VBPendingReboot -ComputerName SERVER01, SERVER02 -IncludeSCCM

    Checks multiple servers including SCCM reboot flags.



    .EXAMPLE

    Get-VBPendingReboot | Where-Object RebootRequired | Select-Object ComputerName, *Pending

    Shows only computers that require a reboot with details.



    .OUTPUTS

    PSCustomObject with properties:

      - ComputerName              : Computer that was checked

      - ComponentBasedServicing   : CBS awaiting reboot

      - WindowsUpdateReboot       : Windows Update requires reboot

      - PendingFileRenameOp       : File rename operations pending

      - ServerFeatureManager      : Server roles/features pending

      - ComputerRenamePending     : Computer rename pending

      - DomainJoinReboot          : Domain join/unjoin pending

      - SCCMReboot                : SCCM client reboot pending (if -IncludeSCCM)

      - RebootRequired            : Overall reboot status

      - LastBootTime              : Last system boot time

      - UptimeDays                : System uptime in days

      - CollectionTime            : When the data was collected

      - Status                    : Success or Failed

      - Error                     : Error message if Status is Failed

    #>



    [CmdletBinding()]

    [OutputType([pscustomobject])]

    param (

        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]

        [Alias('CN', 'Name')]

        [string[]]$ComputerName = $env:COMPUTERNAME,

        

        [switch]$IncludeSCCM,

        

        [PSCredential]$Credential

    )



    begin {

        $scriptBlock = {

            param([bool]$IncludeSCCM)

            

            try {

                # Initialize result object

                $result = [pscustomobject]@{

                    ComputerName            = $env:COMPUTERNAME

                    ComponentBasedServicing = $false

                    WindowsUpdateReboot     = $false

                    PendingFileRenameOp     = $false

                    ServerFeatureManager    = $false

                    ComputerRenamePending   = $false

                    DomainJoinReboot        = $false

                    SCCMReboot              = $false

                    RebootRequired          = $false

                    LastBootTime            = $null

                    UptimeDays              = $null

                    CollectionTime          = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')

                    Status                  = 'Success'

                    Error                   = $null

                }



                # Get system boot time and uptime

                try {

                    $bootTime = (Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop).LastBootUpTime

                    $result.LastBootTime = $bootTime

                    $result.UptimeDays = [math]::Round((Get-Date).Subtract($bootTime).TotalDays, 2)

                } catch {

                    Write-Warning "Could not retrieve boot time: $($_.Exception.Message)"

                }



                # Component-Based Servicing

                $cbsPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'

                $result.ComponentBasedServicing = Test-Path $cbsPath



                # Windows Update Auto Update

                $wuPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'

                $result.WindowsUpdateReboot = Test-Path $wuPath



                # Pending File Rename Operations

                try {

                    $sessionKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager'

                    $pendingRenames = (Get-ItemProperty -Path $sessionKey -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue).PendingFileRenameOperations

                    $result.PendingFileRenameOp = $null -ne $pendingRenames -and $pendingRenames.Count -gt 0

                } catch {

                    $result.PendingFileRenameOp = $false

                }



                # Server Feature Manager (Server 2008+)

                $sfmPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackagesPending'

                if (Test-Path $sfmPath) {

                    $result.ServerFeatureManager = $true

                } else {

                    $sfmPath2 = 'HKLM:\SOFTWARE\Microsoft\ServerManager\CurrentRebootAttempts'

                    $result.ServerFeatureManager = Test-Path $sfmPath2

                }



                # Computer Rename Pending

                try {

                    $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop

                    $result.ComputerRenamePending = $computerSystem.Name -ne $computerSystem.PendingComputerName -and

                                                   ![string]::IsNullOrEmpty($computerSystem.PendingComputerName)

                } catch {

                    $result.ComputerRenamePending = $false

                }



                # Domain Join/Unjoin pending

                $netlogonPath = 'HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon'

                try {

                    $netlogonParams = Get-ItemProperty -Path $netlogonPath -ErrorAction Stop

                    $result.DomainJoinReboot = ($netlogonParams.JoinDomain -eq 1) -or ($netlogonParams.AvoidSpnSet -eq 1)

                } catch {

                    $result.DomainJoinReboot = $false

                }



                # SCCM Client reboot pending (if requested)

                if ($IncludeSCCM) {

                    try {

                        $sccmClient = Get-CimInstance -Namespace 'ROOT\ccm\ClientSDK' -ClassName 'CCM_ClientUtilities' -ErrorAction Stop

                        $rebootStatus = Invoke-CimMethod -InputObject $sccmClient -MethodName 'DetermineIfRebootPending' -ErrorAction Stop

                        $result.SCCMReboot = $rebootStatus.RebootPending -or $rebootStatus.IsHardRebootPending

                    } catch {

                        $result.SCCMReboot = $false

                    }

                }



                # Determine overall reboot requirement

                $rebootChecks = @(

                    $result.ComponentBasedServicing,

                    $result.WindowsUpdateReboot,

                    $result.PendingFileRenameOp,

                    $result.ServerFeatureManager,

                    $result.ComputerRenamePending,

                    $result.DomainJoinReboot

                )

                

                if ($IncludeSCCM) {

                    $rebootChecks += $result.SCCMReboot

                }

                

                $result.RebootRequired = $rebootChecks -contains $true



                return $result



            } catch {

                return [pscustomobject]@{

                    ComputerName            = $env:COMPUTERNAME

                    ComponentBasedServicing = $false

                    WindowsUpdateReboot     = $false

                    PendingFileRenameOp     = $false

                    ServerFeatureManager    = $false

                    ComputerRenamePending   = $false

                    DomainJoinReboot        = $false

                    SCCMReboot              = $false

                    RebootRequired          = $false

                    LastBootTime            = $null

                    UptimeDays              = $null

                    CollectionTime          = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')

                    Status                  = 'Failed'

                    Error                   = $_.Exception.Message

                }

            }

        }

    }



    process {

        foreach ($computer in $ComputerName) {

            try {

                if ($computer -eq $env:COMPUTERNAME -or $computer -eq 'localhost' -or $computer -eq '.') {

                    # Local execution

                    & $scriptBlock $IncludeSCCM.IsPresent

                } else {

                    # Remote execution

                    $invokeParams = @{

                        ComputerName = $computer

                        ScriptBlock  = $scriptBlock

                        ArgumentList = $IncludeSCCM.IsPresent

                        ErrorAction  = 'Stop'

                    }

                    

                    if ($Credential) {

                        $invokeParams.Credential = $Credential

                    }

                    

                    Invoke-Command @invokeParams

                }

            } catch {

                [pscustomobject]@{

                    ComputerName            = $computer

                    ComponentBasedServicing = $false

                    WindowsUpdateReboot     = $false

                    PendingFileRenameOp     = $false

                    ServerFeatureManager    = $false

                    ComputerRenamePending   = $false

                    DomainJoinReboot        = $false

                    SCCMReboot              = $false

                    RebootRequired          = $false

                    LastBootTime            = $null

                    UptimeDays              = $null

                    CollectionTime          = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')

                    Status                  = 'Failed'

                    Error                   = "Connection failed: $($_.Exception.Message)"

                }

            }

        }

    }

}