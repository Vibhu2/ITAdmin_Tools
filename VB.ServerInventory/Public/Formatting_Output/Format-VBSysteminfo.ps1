function Format-VBSystemInfo {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [PSCustomObject[]]$SystemInfo,
        [Parameter()]
        [string]$OutputPath
    )

    begin {
        $output = @()
    }

    process {
        foreach ($info in $SystemInfo) {
            $section = @()
            
            if ($info.Status -eq 'Failed') {
                $section += "ComputerName: $($info.ComputerName)"
                $section += "Error: $($info.Error)"
                $section += ""
            } else {
                $section += "--- COMPUTER IDENTITY ---"
                $section += "ComputerName              : $($info.ComputerName)"
                $section += "DNSHostname               : $($info.DNSHostname)"
                $section += "Domain                    : $($info.Domain)"
                $section += "DomainRole                : $($info.DomainRole)"
                $section += "LogonServer               : $($info.LogonServer)"
                $section += ""
                
                $section += "--- NETWORK / DNS / DHCP ---"
                $section += "PrimaryIP                 : $($info.PrimaryIP)"
                $section += "DHCPStatus                : $($info.DHCPStatus)"
                $section += "DHCPServer                : $($info.DHCPServer)"
                $section += "DHCPLeaseObtained         : $($info.DHCPLeaseObtained)"
                $section += "DHCPLeaseExpires          : $($info.DHCPLeaseExpires)"
                $section += "DNSServers                : $($info.DNSServers)"
                $section += "DNSSuffix                 : $($info.DNSSuffix)"
                $section += ""
                
                $section += "--- OPERATING SYSTEM ---"
                $section += "OSName                    : $($info.OSName)"
                $section += "OSVersion                 : $($info.OSVersion)"
                $section += "OSBuild                   : $($info.OSBuild)"
                $section += "OSDisplayVersion          : $($info.OSDisplayVersion)"
                $section += "OSInstallDate             : $($info.OSInstallDate)"
                $section += ""
                
                $section += "--- HARDWARE ---"
                $section += "SystemManufacturer        : $($info.SystemManufacturer)"
                $section += "SystemModel               : $($info.SystemModel)"
                $section += "Processor                 : $($info.Processor)"
                $section += "Cores                     : $($info.Cores)"
                $section += "LogicalProcessors         : $($info.LogicalProcessors)"
                $section += "TotalMemoryGB             : $($info.TotalMemoryGB)"
                $section += "FreeMemoryGB              : $($info.FreeMemoryGB)"
                $section += ""
                
                $section += "--- BIOS/FIRMWARE ---"
                $section += "BIOSType                  : $($info.BIOSType)"
                $section += "BIOSVersion               : $($info.BIOSVersion)"
                $section += "BIOSManufacturer          : $($info.BIOSManufacturer)"
                $section += "BIOSSerialNumber          : $($info.BIOSSerialNumber)"
                $section += "BIOSReleaseDate           : $($info.BIOSReleaseDate)"
                $section += ""
                
                $section += "--- SYSTEM ---"
                $section += "TimeZone                  : $($info.TimeZone)"
                $section += "LastBootTime              : $($info.LastBootTime)"
                $section += "Status                    : $($info.Status)"
                $section += "CollectionTime            : $($info.CollectionTime)"
                $section += ""
                $section += "=" * 80
                $section += ""
            }

            if ($OutputPath) {
                $output += $section
            } else {
                foreach ($line in $section) {
                    if ($line -match "^---") {
                        Write-Host $line -ForegroundColor Cyan
                    } elseif ($line -match "^Error:") {
                        Write-Host $line -ForegroundColor Red
                    } else {
                        Write-Host $line
                    }
                }
            }
        }
    }

    end {
        if ($OutputPath) {
            try {
                $output | Out-File -FilePath $OutputPath -Encoding UTF8 -Force
                Write-Host "Report saved to: $OutputPath" -ForegroundColor Green
            }
            catch {
                Write-Error "Failed to save report: $_"
            }
        }
    }
}

Get-VBSystemInfo | Format-VBSystemInfo
Get-VBSystemInfo | Format-VBSystemInfo -OutputPath "C:\Reports\SystemInfo.txt"