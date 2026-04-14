# ============================================================
# FUNCTION : Get-VBRDSUserInformation
# VERSION  : 1.0.0
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu
# PURPOSE  : Collect RDS user session information
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Collects active RDS user session information from local or remote Remote Desktop Session Host servers.

.DESCRIPTION
    Get-VBRDSUserInformation retrieves detailed information about active RDS user sessions from
    local or remote Remote Desktop Session Host servers. Returns structured objects containing
    session details, user information, and connection metadata.

.PARAMETER ComputerName
    Target RDS server(s). Defaults to local machine. Accepts pipeline input.
    Supports multiple servers and fully-qualified domain names.

.PARAMETER Credential
    Alternate credentials for remote RDS server access.

.EXAMPLE
    Get-VBRDSUserInformation
    Retrieves active RDS sessions from the local server.

.EXAMPLE
    Get-VBRDSUserInformation -ComputerName RDS-HOST01
    Retrieves active RDS sessions from remote server RDS-HOST01.

.EXAMPLE
    'RDS-HOST01','RDS-HOST02' | Get-VBRDSUserInformation -Credential (Get-Credential)
    Retrieves active RDS sessions from multiple servers using pipeline input.

.OUTPUTS
    [PSCustomObject]: ComputerName, UserName, SessionId, SessionState, HostServer, ClientName,
    ClientIP, LogonTime, Status, CollectionTime

.NOTES
    Version  : 1.0.0
    Author   : Vibhu
    Modified : 10-04-2026
    Category : Services
#>

function Get-VBRDSUserInformation {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Name', 'Server', 'Host')]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        [PSCredential]$Credential
    )

    process {
        foreach ($computer in $ComputerName) {
            try {
                # Step 1 -- Define the RDS session collection script block
                $ScriptBlock = {
                    Get-RDUserSession -ErrorAction SilentlyContinue | Select-Object -Property UserName, SessionId, SessionState, HostServer, ClientName, ClientIP, LogonTime
                }

                # Step 2 -- Determine if local or remote execution
                if ($computer -eq $env:COMPUTERNAME) {
                    # Step 3 -- Execute locally
                    $UserSessions = & $ScriptBlock
                } else {
                    # Step 4 -- Execute remotely
                    $UserSessions = Invoke-Command -ComputerName $computer -Credential $Credential -ScriptBlock $ScriptBlock
                }

                # Step 5 -- Return each session as a separate object
                if ($UserSessions) {
                    foreach ($session in $UserSessions) {
                        [PSCustomObject]@{
                            ComputerName   = $computer
                            UserName       = $session.UserName
                            SessionId      = $session.SessionId
                            SessionState   = $session.SessionState
                            HostServer     = $session.HostServer
                            ClientName     = $session.ClientName
                            ClientIP       = $session.ClientIP
                            LogonTime      = $session.LogonTime
                            Status         = 'Success'
                            CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                        }
                    }
                } else {
                    # Step 6 -- No sessions found but no error
                    [PSCustomObject]@{
                        ComputerName   = $computer
                        UserName       = $null
                        SessionId      = $null
                        SessionState   = $null
                        HostServer     = $null
                        ClientName     = $null
                        ClientIP       = $null
                        LogonTime      = $null
                        Status         = 'Success'
                        CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                    }
                }
            }
            catch {
                # Step 7 -- Return error object
                [PSCustomObject]@{
                    ComputerName   = $computer
                    UserName       = $null
                    SessionId      = $null
                    SessionState   = $null
                    HostServer     = $null
                    ClientName     = $null
                    ClientIP       = $null
                    LogonTime      = $null
                    Error          = $_.Exception.Message
                    Status         = 'Failed'
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}
