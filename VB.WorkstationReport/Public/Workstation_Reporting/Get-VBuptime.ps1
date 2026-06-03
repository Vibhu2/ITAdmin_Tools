# ============================================================
# SCRIPT  : Get-VBUptime.ps1
# VERSION : 1.0.0
# CHANGED : 03-06-2026 -- Initial release
# AUTHOR  : Vibhu
# PURPOSE : Returns system uptime for one or more computers
# ENCODING: UTF-8 with BOM -- do not re-save without BOM
# ------------------------------------------------------------
# CHANGELOG (last 3-5 only -- full history in Git)
# v1.0.0 -- 03-06-2026 -- Initial release
# ============================================================

function Get-VBUptime {
    <#
    .SYNOPSIS
        Returns system uptime for one or more computers.
    .DESCRIPTION
        Queries Win32_OperatingSystem via CIM and returns uptime
        broken into Days, Hours, Minutes, and Seconds, along with
        the raw LastBootUpTime timestamp. Runs locally without WinRM
        when targeting the local machine; uses DCOM for remote targets.
    .PARAMETER ComputerName
        Target computer(s). Defaults to the local machine.
    .EXAMPLE
        Get-VBUptime
    .EXAMPLE
        Get-VBUptime -ComputerName SRV01, SRV02
    .EXAMPLE
        'SRV01','SRV02' | Get-VBUptime
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string[]]$ComputerName = $env:COMPUTERNAME
    )
    process {
        foreach ($computer in $ComputerName) {
            try {
                if ($computer -eq $env:COMPUTERNAME) {
                    $os = Get-CimInstance -ClassName Win32_OperatingSystem
                } else {
                    $sessionOption = New-CimSessionOption -Protocol Dcom
                    $session = New-CimSession -ComputerName $computer -SessionOption $sessionOption
                    $os = Get-CimInstance -ClassName Win32_OperatingSystem -CimSession $session
                    Remove-CimSession -CimSession $session
                }
                $uptime = (Get-Date) - $os.LastBootUpTime

                [PSCustomObject]@{
                    ComputerName   = $computer
                    Days           = $uptime.Days
                    Hours          = $uptime.Hours
                    Minutes        = $uptime.Minutes
                    Seconds        = $uptime.Seconds
                    LastBootUpTime = $os.LastBootUpTime
                    Error          = $null
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
            catch {
                [PSCustomObject]@{
                    ComputerName   = $computer
                    Days           = $null
                    Hours          = $null
                    Minutes        = $null
                    Seconds        = $null
                    LastBootUpTime = $null
                    Error          = $_.Exception.Message
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}
