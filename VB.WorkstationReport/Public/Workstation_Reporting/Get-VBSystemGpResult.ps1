#Requires -Version 5.1
# ============================================================
# FUNCTION : Get-VBSystemGpResult
# VERSION  : 1.0.0
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Parse gpresult /r /scope computer into PSCustomObjects
# ENCODING : UTF-8 with BOM
# ============================================================

<#
REMOTING PREREQUISITES:
  Target: WinRM running, Enable-PSRemoting -Force, port 5985 open
  Domain: current token used if -Credential omitted
  Workgroup: requires -Credential or TrustedHosts configured
  Production: use HTTPS (port 5986) with valid certificate
#>

function Get-VBSystemGpResult {
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
                Write-Verbose "[$computer] Capturing gpresult /r /scope computer"

                $credSplat = if ($Credential -ne [PSCredential]::Empty) { @{ Credential = $Credential } } else { @{} }

                $scriptBlock = {
                    $GpOutput = gpresult /r /scope computer 2>$null
                    if ($LASTEXITCODE -ne 0) {
                        throw "gpresult exited with code $LASTEXITCODE"
                    }
                    $GpOutput | Where-Object { $_ -is [string] }
                }

                if ($computer -eq $env:COMPUTERNAME) {
                    $Lines = & $scriptBlock
                } else {
                    $Lines = Invoke-Command -ComputerName $computer @credSplat -ScriptBlock $scriptBlock
                }

                # Step 1 -- Helper: extract single field value, returns $null if not found
                $fieldValue = {
                    param([string[]]$L, [string]$Pattern)
                    $m = $L | Select-String -Pattern $Pattern | Select-Object -First 1
                    if ($null -eq $m) { return $null }
                    $m.Matches[0].Groups[1].Value.Trim()
                }

                # Step 2 -- Parse system info block
                $OsConfig      = & $fieldValue -L $Lines -Pattern 'OS Configuration:\s+(.+)'
                $OsVersion     = & $fieldValue -L $Lines -Pattern 'OS Version:\s+(.+)'
                $SiteName      = & $fieldValue -L $Lines -Pattern 'Site Name:\s+(.+)'
                $SlowLink      = & $fieldValue -L $Lines -Pattern 'Connected over a slow link\?:\s+(.+)'
                $LastApplied   = & $fieldValue -L $Lines -Pattern 'Last time Group Policy was applied:\s+(.+)'
                $AppliedFrom   = & $fieldValue -L $Lines -Pattern 'Group Policy was applied from:\s+(.+)'
                $SlowThreshold = & $fieldValue -L $Lines -Pattern 'Group Policy slow link threshold:\s+(.+)'
                $DomainName    = & $fieldValue -L $Lines -Pattern 'Domain Name:\s+(.+)'
                $DomainType    = & $fieldValue -L $Lines -Pattern 'Domain Type:\s+(.+)'

                # Step 3 -- Parse Applied GPOs block
                $GpoStartLine = ($Lines | Select-String 'Applied Group Policy Objects' | Select-Object -First 1).LineNumber
                $GpoEndLine   = ($Lines | Select-String 'The computer is a part of' | Select-Object -First 1).LineNumber

                $AppliedGPOs = $Lines[($GpoStartLine)..($GpoEndLine - 2)] |
                    Where-Object { $_.Trim() -ne '' -and $_ -notmatch '---' -and $_ -notmatch 'Applied Group Policy' } |
                    ForEach-Object {
                        [PSCustomObject]@{
                            ComputerName = $computer
                            GPOName      = $_.Trim()
                        }
                    }

                # Step 4 -- Parse Security Groups block
                $SgStartLine = ($Lines | Select-String 'The computer is a part of the following security groups' | Select-Object -First 1).LineNumber

                $SecurityGroups = $Lines[($SgStartLine)..($Lines.Count - 1)] |
                    Where-Object { $_.Trim() -ne '' -and $_ -notmatch '---' -and $_ -notmatch 'security groups' } |
                    ForEach-Object {
                        [PSCustomObject]@{
                            ComputerName  = $computer
                            SecurityGroup = $_.Trim()
                        }
                    }

                # Step 5 -- Build result object
                [PSCustomObject]@{
                    ComputerName      = $computer
                    Status            = 'Success'
                    OsConfiguration   = $OsConfig
                    OsVersion         = $OsVersion
                    SiteName          = $SiteName
                    SlowLink          = $SlowLink
                    LastApplied       = $LastApplied
                    AppliedFrom       = $AppliedFrom
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
                Write-Warning "[$computer] $($_.Exception.Message)"
                [PSCustomObject]@{
                    ComputerName      = $computer
                    Status            = 'Failed'
                    OsConfiguration   = $null
                    OsVersion         = $null
                    SiteName          = $null
                    SlowLink          = $null
                    LastApplied       = $null
                    AppliedFrom       = $null
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
