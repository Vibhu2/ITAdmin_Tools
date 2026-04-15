# ============================================================
# FUNCTION : Get-VBAdAccountWithNoLogin
# VERSION  : 1.0.0
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : VB (Vibhu Bhatnagar)
# PURPOSE  : Find enabled AD user accounts with no logon history
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Returns enabled Active Directory user accounts that have never logged on.

.DESCRIPTION
    Queries Active Directory for enabled user accounts where lastLogonTimestamp
    has never been set, indicating the account has never been used to log on.
    Useful for identifying stale or provisioned-but-unused accounts during
    AD hygiene audits.

    Returns one PSCustomObject per matching user. Supports targeting a specific
    Domain Controller via -ComputerName and alternate credentials via -Credential.

.PARAMETER ComputerName
    Domain Controller to query. Defaults to the local machine.
    Accepts pipeline input and multiple values.

.PARAMETER Credential
    Alternate credentials for the AD query. Optional on domain-joined machines.

.EXAMPLE
    Get-VBAdAccountWithNoLogin

    Returns all never-logged-on enabled accounts from the default DC.

.EXAMPLE
    Get-VBAdAccountWithNoLogin -ComputerName DC01

    Runs the query against DC01 specifically.

.EXAMPLE
    Get-VBAdAccountWithNoLogin -ComputerName DC01 -Credential (Get-Credential)

    Runs the query against DC01 using alternate credentials.

.EXAMPLE
    Get-VBAdAccountWithNoLogin | Sort-Object WhenCreated | Format-Table -AutoSize

    Returns results sorted by creation date for review.

.OUTPUTS
    [PSCustomObject]: ComputerName, Name, SamAccountName, Enabled, WhenCreated, Status, CollectionTime

.NOTES
    Version  : 1.0.0
    Author   : Vibhu Bhatnagar
    Modified : 10-04-2026
    Category : AD User Health
#>
function Get-VBAdAccountWithNoLogin {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Name', 'Server', 'Host')]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        [PSCredential]$Credential
    )
    begin {
        Import-Module ActiveDirectory -ErrorAction Stop
    }


    process {
        foreach ($computer in $ComputerName) {
            try {
                # Step 1 -- Build AD query parameters
                $AdParams = @{
                    Filter     = { Enabled -eq $true -and lastLogonTimestamp -notlike '*' }
                    Properties = 'lastLogonTimestamp', 'whenCreated'
                }
                if ($computer -ne $env:COMPUTERNAME) { $AdParams['Server'] = $computer }
                if ($Credential) { $AdParams['Credential'] = $Credential }

                # Step 2 -- Query AD and emit one object per matching user
                $Users = Get-ADUser @AdParams

                foreach ($user in $Users) {
                    [PSCustomObject]@{
                        ComputerName   = $computer
                        Name           = $user.Name
                        SamAccountName = $user.SamAccountName
                        Enabled        = $user.Enabled
                        WhenCreated    = $user.whenCreated
                        Status         = 'Success'
                        CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                    }
                }
            }
            catch {
                [PSCustomObject]@{
                    ComputerName = $computer
                    Error        = $_.Exception.Message
                    Status       = 'Failed'
                }
            }
        }
    }
}
