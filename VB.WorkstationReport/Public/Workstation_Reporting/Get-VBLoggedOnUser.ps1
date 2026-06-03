#Requires -Version 5.1
# ============================================================
# FUNCTION : Get-VBLoggedOnUser
# MODULE   : VB.WorkstationReport
# VERSION  : 1.0.0
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Returns interactive logon sessions with quser session detail
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Returns interactive logon sessions on local or remote computers.

.DESCRIPTION
    Combines Win32_LogonSession (LogonType 2 and 10) with quser session data
    to return one object per interactive session with user, domain, SID,
    logon type, session name, session ID, state, idle time, and logon time.

    quser enrichment is best-effort -- if quser returns no match for a session,
    session-specific fields are $null rather than failing the entire result.

    Supports local and remote execution. Credentials are only passed to
    CIM and Invoke-Command when explicitly supplied.

.PARAMETER ComputerName
    Target computer(s). Defaults to local machine. Accepts pipeline input.
    Aliases: CN, MachineName, PSComputerName

.PARAMETER Credential
    Alternate credentials for remote execution.

.EXAMPLE
    Get-VBLoggedOnUser
    Returns all interactive sessions on the local computer.

.EXAMPLE
    Get-VBLoggedOnUser -ComputerName WS001, WS002
    Returns interactive sessions from two remote workstations.

.EXAMPLE
    'WS001','WS002' | Get-VBLoggedOnUser -Credential $cred
    Pipeline input with alternate credentials.

.OUTPUTS
    PSCustomObject with: ComputerName, Status, Name, Domain, SID,
    LogonType, LogonId, Session, SessionId, State, IdleTime, LogonTime,
    Error, CollectionTime

.NOTES
    Version      : 1.0.0
    Author       : Vibhu Bhatnagar
    Category     : User Management
    Requirements :
      - PowerShell 5.1 or higher
      - quser.exe available on target (standard on all Windows desktop/server SKUs)
      - PowerShell Remoting enabled for remote targets

REMOTING PREREQUISITES:
  Target: WinRM running, Enable-PSRemoting -Force, port 5985 open
  Domain: current token used if -Credential omitted
  Workgroup: requires -Credential or TrustedHosts configured
  Production: use HTTPS (port 5986) with valid certificate
#>
function Get-VBLoggedOnUser {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('CN', 'MachineName', 'PSComputerName')]
        [string[]]$ComputerName = $env:COMPUTERNAME,

        [Parameter()]
        [PSCredential]$Credential
    )

    process {
        foreach ($computer in $ComputerName) {
            try {
                Write-Verbose "[$computer] Collecting interactive logon sessions"

                $credSplat = if ($Credential) { @{ Credential = $Credential } } else { @{} }

                # Step 1 -- Collect quser and CIM data: local or remote
                if ($computer -eq $env:COMPUTERNAME) {

                    # Step 1a -- Parse quser locally
                    $QuserSessions = quser 2>$null |
                        Select-Object -Skip 1 |
                        ForEach-Object {
                            if ($_ -match '^\s*>?(?<User>\S+)\s+(?<Session>\S+)\s+(?<Id>\d+)\s+(?<State>\w+)\s+(?<Idle>\S+)\s+(?<LogonTime>.+)$') {
                                [PSCustomObject]@{
                                    UserName  = $Matches.User
                                    Session   = $Matches.Session
                                    SessionId = [int]$Matches.Id
                                    State     = $Matches.State
                                    IdleTime  = $Matches.Idle
                                    LogonTime = $Matches.LogonTime.Trim()
                                }
                            }
                        }

                    # Step 1b -- Query CIM locally
                    $LogonSessions = Get-CimInstance -ClassName Win32_LogonSession -ErrorAction Stop |
                        Where-Object { $_.LogonType -in 2, 10 }

                    $Results = foreach ($session in $LogonSessions) {
                        Get-CimAssociatedInstance -InputObject $session -Association Win32_LoggedOnUser |
                            ForEach-Object {
                                $user  = $_
                                $match = $QuserSessions |
                                    Where-Object { $_.UserName -eq $user.Name } |
                                    Select-Object -First 1

                                [PSCustomObject]@{
                                    Name      = $user.Name
                                    Domain    = $user.Domain
                                    SID       = $user.SID
                                    LogonType = $session.LogonType
                                    LogonId   = $session.LogonId
                                    Session   = $match.Session
                                    SessionId = $match.SessionId
                                    State     = $match.State
                                    IdleTime  = $match.IdleTime
                                    LogonTime = $match.LogonTime
                                }
                            }
                    }

                } else {

                    # Step 1c -- Run quser and CIM remotely
                    $Results = Invoke-Command -ComputerName $computer @credSplat -ScriptBlock {

                        $QuserSessions = quser 2>$null |
                            Select-Object -Skip 1 |
                            ForEach-Object {
                                if ($_ -match '^\s*>?(?<User>\S+)\s+(?<Session>\S+)\s+(?<Id>\d+)\s+(?<State>\w+)\s+(?<Idle>\S+)\s+(?<LogonTime>.+)$') {
                                    [PSCustomObject]@{
                                        UserName  = $Matches.User
                                        Session   = $Matches.Session
                                        SessionId = [int]$Matches.Id
                                        State     = $Matches.State
                                        IdleTime  = $Matches.Idle
                                        LogonTime = $Matches.LogonTime.Trim()
                                    }
                                }
                            }

                        $LogonSessions = Get-CimInstance -ClassName Win32_LogonSession -ErrorAction Stop |
                            Where-Object { $_.LogonType -in 2, 10 }

                        foreach ($session in $LogonSessions) {
                            Get-CimAssociatedInstance -InputObject $session -Association Win32_LoggedOnUser |
                                ForEach-Object {
                                    $user  = $_
                                    $match = $QuserSessions |
                                        Where-Object { $_.UserName -eq $user.Name } |
                                        Select-Object -First 1

                                    [PSCustomObject]@{
                                        Name      = $user.Name
                                        Domain    = $user.Domain
                                        SID       = $user.SID
                                        LogonType = $session.LogonType
                                        LogonId   = $session.LogonId
                                        Session   = $match.Session
                                        SessionId = $match.SessionId
                                        State     = $match.State
                                        IdleTime  = $match.IdleTime
                                        LogonTime = $match.LogonTime
                                    }
                                }
                        }
                    }
                }

                # Step 2 -- Deduplicate and emit result objects
                $CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')

                $Results |
                    Sort-Object Domain, Name -Unique |
                    ForEach-Object {
                        [PSCustomObject]@{
                            ComputerName = $computer
                            Status       = 'Success'
                            Name         = $_.Name
                            Domain       = $_.Domain
                            SID          = $_.SID
                            LogonType    = $_.LogonType
                            LogonId      = $_.LogonId
                            Session      = $_.Session
                            SessionId    = $_.SessionId
                            State        = $_.State
                            IdleTime     = $_.IdleTime
                            LogonTime    = $_.LogonTime
                            Error        = $null
                            CollectionTime = $CollectionTime
                        }
                    }
            }
            catch {
                Write-Warning "[$computer] $($_.Exception.Message)"
                [PSCustomObject]@{
                    ComputerName   = $computer
                    Status         = 'Failed'
                    Name           = $null
                    Domain         = $null
                    SID            = $null
                    LogonType      = $null
                    LogonId        = $null
                    Session        = $null
                    SessionId      = $null
                    State          = $null
                    IdleTime       = $null
                    LogonTime      = $null
                    Error          = $_.Exception.Message
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}
