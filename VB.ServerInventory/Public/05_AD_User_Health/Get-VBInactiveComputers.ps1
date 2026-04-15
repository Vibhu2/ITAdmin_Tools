# ============================================================
# FUNCTION : Get-VBInactiveComputers
# VERSION  : 1.0.0
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Get Active Directory computers inactive for 90+ days
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Get Active Directory computers inactive for 90 days or more.
.DESCRIPTION
    Queries Active Directory for enabled computers whose lastLogonTimestamp is
    older than 90 days. Returns computer details including DNS hostname, last
    logon date, and resolved IP address (if available).
.PARAMETER ComputerName
    Domain Controller to query. Defaults to local machine. Accepts pipeline input.
.PARAMETER Credential
    Alternate credentials for the AD query.
.EXAMPLE
    Get-VBInactiveComputers
.EXAMPLE
    Get-VBInactiveComputers -ComputerName DC01
.EXAMPLE
    'DC01' | Get-VBInactiveComputers -Credential (Get-Credential)
.OUTPUTS
    [PSCustomObject]: ComputerName, Name, DNSHostName, LastLogonDate, IPAddress, Status, CollectionTime
.NOTES
    Version  : 1.0.0
    Author   : Vibhu Bhatnagar
    Modified : 10-04-2026
    Category : AD User Health
#>

function Get-VBInactiveComputers {
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
                # Step 1 -- Define the inactivity threshold (90 days ago)
                $thresholdDate = (Get-Date).AddDays(-90)

                # Step 2 -- Build AD query parameters
                $AdParams = @{
                    Filter     = "enabled -eq `$true -and lastLogonTimestamp -lt `$($thresholdDate.FileTime)"
                    Properties = 'lastLogonTimestamp', 'DNSHostName'
                }
                if ($computer -ne $env:COMPUTERNAME) {
                    $AdParams['Server'] = $computer
                }
                if ($Credential) {
                    $AdParams['Credential'] = $Credential
                }

                # Step 3 -- Get inactive computers
                $inactiveComputers = Get-ADComputer @AdParams

                # Step 4 -- Emit results
                foreach ($comp in $inactiveComputers) {
                    $lastLogonDate = if ($comp.lastLogonTimestamp) {
                        [DateTime]::FromFileTime($comp.lastLogonTimestamp)
                    }
                    else {
                        $null
                    }

                    # Step 5 -- Resolve IP address if DNS hostname exists
                    $ipAddress = 'Unresolved'
                    if ($comp.DNSHostName) {
                        try {
                            $dnsRecord = Resolve-DnsName -Name $comp.DNSHostName -ErrorAction Stop | Where-Object { $_.Type -eq 'A' }
                            if ($dnsRecord) {
                                $ipAddress = $dnsRecord[0].IPAddress
                            }
                        }
                        catch {
                            $ipAddress = 'Unresolved'
                        }
                    }
                    else {
                        $ipAddress = 'No DNSHostName'
                    }

                    [PSCustomObject]@{
                        ComputerName   = $computer
                        Name           = $comp.Name
                        DNSHostName    = $comp.DNSHostName
                        LastLogonDate  = $lastLogonDate
                        IPAddress      = $ipAddress
                        Status         = 'Success'
                        CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                    }
                }
            }
            catch {
                [PSCustomObject]@{
                    ComputerName   = $computer
                    Name           = $null
                    DNSHostName    = $null
                    LastLogonDate  = $null
                    IPAddress      = $null
                    Error          = $_.Exception.Message
                    Status         = 'Failed'
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}
