#Requires -Version 5.1
# ============================================================
# FUNCTION : Get-VBLoggedOnUserGpResult
# MODULE   : VB.WorkstationReport
# VERSION  : 1.0.0
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Parse gpresult /r /scope user for each logged-on user into PSCustomObjects
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Returns user-scoped Group Policy results for every logged-on interactive user.

.DESCRIPTION
    Calls Get-VBLoggedOnUser to discover all interactive sessions, then runs
    gpresult /r /scope user /user <name> for each user found. Returns one result
    object per user per computer.

    If no users are logged on, a single result object is returned with
    Status 'NoUsers' and all data properties as $null.

    Supports local and remote execution. Credentials are passed to both
    Get-VBLoggedOnUser and Invoke-Command when supplied.

.PARAMETER ComputerName
    Target computer(s). Defaults to local machine. Accepts pipeline input.
    Aliases: CN, MachineName, PSComputerName

.PARAMETER Credential
    Alternate credentials for remote execution.

.EXAMPLE
    Get-VBLoggedOnUserGpResult
    Returns user GP results for all logged-on users on the local computer.

.EXAMPLE
    Get-VBLoggedOnUserGpResult -ComputerName WS001, WS002
    Returns user GP results from two remote workstations.

.EXAMPLE
    'WS001','WS002' | Get-VBLoggedOnUserGpResult -Credential $cred
    Pipeline input with alternate credentials.

.OUTPUTS
    PSCustomObject with: ComputerName, Status, UserName, LastApplied,
    AppliedFrom, SlowLink, SlowLinkThreshold, DomainName, DomainType,
    AppliedGPOs, SecurityGroups, Error, CollectionTime

.NOTES
    Version      : 1.0.0
    Author       : Vibhu Bhatnagar
    Category     : Group Policy
    Requirements :
      - PowerShell 5.1 or higher
      - gpresult.exe available on target (standard on all Windows SKUs)
      - PowerShell Remoting enabled for remote targets
      - Get-VBLoggedOnUser available in the same session

REMOTING PREREQUISITES:
  Target: WinRM running, Enable-PSRemoting -Force, port 5985 open
  Domain: current token used if -Credential omitted
  Workgroup: requires -Credential or TrustedHosts configured
  Production: use HTTPS (port 5986) with valid certificate
#>
function Get-VBLoggedOnUserGpResult {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('CN', 'MachineName', 'PSComputerName')]
        [string[]]$ComputerName = $env:COMPUTERNAME,

        [Parameter()]
        [PSCredential]$Credential = [PSCredential]::Empty
    )

    process {
        foreach ($computer in $ComputerName) {
            try {
                Write-Verbose "[$computer] Discovering logged-on users"

                $credSplat = if ($Credential -ne [PSCredential]::Empty) { @{ Credential = $Credential } } else { @{} }

                # Step 1 -- Discover logged-on interactive users via Get-VBLoggedOnUser
                $loggedOnParams = @{ ComputerName = $computer }
                if ($Credential -ne [PSCredential]::Empty) { $loggedOnParams['Credential'] = $Credential }

                $LoggedOnUsers = Get-VBLoggedOnUser @loggedOnParams |
                    Where-Object { $_.Status -eq 'Success' -and $null -ne $_.Name }

                if ($null -eq $LoggedOnUsers -or @($LoggedOnUsers).Count -eq 0) {
                    Write-Verbose "[$computer] No interactive users found"
                    [PSCustomObject]@{
                        ComputerName      = $computer
                        Status            = 'NoUsers'
                        UserName          = $null
                        LastApplied       = $null
                        AppliedFrom       = $null
                        SlowLink          = $null
                        SlowLinkThreshold = $null
                        DomainName        = $null
                        DomainType        = $null
                        AppliedGPOs       = $null
                        SecurityGroups    = $null
                        Error             = $null
                        CollectionTime    = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                    }
                    continue
                }

                # Step 2 -- Run gpresult for each user
                foreach ($user in $LoggedOnUsers) {
                    $userName = $user.Name
                    Write-Verbose "[$computer] Capturing gpresult /r /scope user /user $userName"

                    try {
                        $captureBlock = {
                            param([string]$TargetUser)
                            $GpOutput = gpresult /r /scope user /user $TargetUser 2>$null
                            if ($LASTEXITCODE -ne 0) {
                                throw "gpresult exited with code $LASTEXITCODE for user $TargetUser"
                            }
                            $GpOutput | Where-Object { $_ -is [string] }
                        }

                        if ($computer -eq $env:COMPUTERNAME) {
                            $Lines = & $captureBlock -TargetUser $userName
                        } else {
                            $Lines = Invoke-Command -ComputerName $computer @credSplat -ScriptBlock $captureBlock -ArgumentList $userName
                        }

                        # Step 3 -- Helper: extract single field value, returns $null if not found
                        $fieldValue = {
                            param([string[]]$L, [string]$Pattern)
                            $m = $L | Select-String -Pattern $Pattern | Select-Object -First 1
                            if ($null -eq $m) { return $null }
                            $m.Matches[0].Groups[1].Value.Trim()
                        }

                        # Step 4 -- Parse user info block
                        $SlowLink      = & $fieldValue -L $Lines -Pattern 'Connected over a slow link\?:\s+(.+)'
                        $LastApplied   = & $fieldValue -L $Lines -Pattern 'Last time Group Policy was applied:\s+(.+)'
                        $AppliedFrom   = & $fieldValue -L $Lines -Pattern 'Group Policy was applied from:\s+(.+)'
                        $SlowThreshold = & $fieldValue -L $Lines -Pattern 'Group Policy slow link threshold:\s+(.+)'
                        $DomainName    = & $fieldValue -L $Lines -Pattern 'Domain Name:\s+(.+)'
                        $DomainType    = & $fieldValue -L $Lines -Pattern 'Domain Type:\s+(.+)'

                        # Step 5 -- Parse Applied GPOs block
                        $GpoStartLine = ($Lines | Select-String 'Applied Group Policy Objects' | Select-Object -First 1).LineNumber
                        $GpoEndLine   = ($Lines | Select-String 'The user is a part of' | Select-Object -First 1).LineNumber

                        $AppliedGPOs = $Lines[($GpoStartLine)..($GpoEndLine - 2)] |
                            Where-Object { $_.Trim() -ne '' -and $_ -notmatch '---' -and $_ -notmatch 'Applied Group Policy' } |
                            ForEach-Object {
                                [PSCustomObject]@{
                                    ComputerName = $computer
                                    UserName     = $userName
                                    GPOName      = $_.Trim()
                                }
                            }

                        # Step 6 -- Parse Security Groups block
                        $SgStartLine = ($Lines | Select-String 'The user is a part of the following security groups' | Select-Object -First 1).LineNumber

                        $SecurityGroups = $Lines[($SgStartLine)..($Lines.Count - 1)] |
                            Where-Object { $_.Trim() -ne '' -and $_ -notmatch '---' -and $_ -notmatch 'security groups' } |
                            ForEach-Object {
                                [PSCustomObject]@{
                                    ComputerName  = $computer
                                    UserName      = $userName
                                    SecurityGroup = $_.Trim()
                                }
                            }

                        # Step 7 -- Emit result object for this user
                        [PSCustomObject]@{
                            ComputerName      = $computer
                            Status            = 'Success'
                            UserName          = $userName
                            LastApplied       = $LastApplied
                            AppliedFrom       = $AppliedFrom
                            SlowLink          = $SlowLink
                            SlowLinkThreshold = $SlowThreshold
                            DomainName        = $DomainName
                            DomainType        = $DomainType
                            AppliedGPOs       = $AppliedGPOs
                            SecurityGroups    = $SecurityGroups
                            Error             = $null
                            CollectionTime    = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                        }
                    }
                    catch {
                        # Per-user failure -- emit failure object and continue to next user
                        Write-Warning "[$computer] $userName -- $($_.Exception.Message)"
                        [PSCustomObject]@{
                            ComputerName      = $computer
                            Status            = 'Failed'
                            UserName          = $userName
                            LastApplied       = $null
                            AppliedFrom       = $null
                            SlowLink          = $null
                            SlowLinkThreshold = $null
                            DomainName        = $null
                            DomainType        = $null
                            AppliedGPOs       = $null
                            SecurityGroups    = $null
                            Error             = $_.Exception.Message
                            CollectionTime    = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                        }
                    }
                }
            }
            catch {
                # Computer-level failure -- single failure object for the whole target
                Write-Warning "[$computer] $($_.Exception.Message)"
                [PSCustomObject]@{
                    ComputerName      = $computer
                    Status            = 'Failed'
                    UserName          = $null
                    LastApplied       = $null
                    AppliedFrom       = $null
                    SlowLink          = $null
                    SlowLinkThreshold = $null
                    DomainName        = $null
                    DomainType        = $null
                    AppliedGPOs       = $null
                    SecurityGroups    = $null
                    Error             = $_.Exception.Message
                    CollectionTime    = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}
