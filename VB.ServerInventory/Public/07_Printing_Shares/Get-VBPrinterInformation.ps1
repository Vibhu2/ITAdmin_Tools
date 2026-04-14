# ============================================================
# FUNCTION : Get-VBPrinterInformation
# VERSION  : 1.0.0
# CHANGED  : 10-04-2026 -- Initial VB-compliant release
# AUTHOR   : Vibhu
# PURPOSE  : Enumerate printers and driver information on local and remote systems
# ENCODING : UTF-8 with BOM
# ============================================================

<#
.SYNOPSIS
    Enumerate printers and driver information on local and remote systems.

.DESCRIPTION
    Retrieves detailed printer and driver information from Windows systems.
    Correlates printer data with driver information for complete hardware inventory.

.PARAMETER ComputerName
    Target computer(s). Defaults to local machine. Accepts pipeline input.
    Supports aliases: Name, Server, Host.

.PARAMETER Credential
    Alternate credentials for remote execution.

.EXAMPLE
    Get-VBPrinterInformation

.EXAMPLE
    Get-VBPrinterInformation -ComputerName SERVER01

.EXAMPLE
    'SRV01','SRV02' | Get-VBPrinterInformation -Credential (Get-Credential)

.OUTPUTS
    [PSCustomObject]: ComputerName, PrinterName, DriverName, PortName, Published, Shared, ShareName, Type, DeviceType, Manufacturer, InfPath, Status, CollectionTime

.NOTES
    Version  : 1.0.0
    Author   : Vibhu
    Modified : 10-04-2026
    Category : Printing
#>

function Get-VBPrinterInformation {
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
                # Step 1 -- Define script block for printer enumeration
                $scriptBlock = {
                    $printers = Get-Printer |
                        Select-Object Name, DriverName, PortName, Published, Shared, ShareName, Type, DeviceType

                    $drivers = Get-PrinterDriver |
                        Select-Object Name, Manufacturer, InfPath

                    # Step 2 -- Correlate printer with driver information
                    $combined = $printers | ForEach-Object {
                        $printer = $_
                        $driver = $drivers | Where-Object { $_.Name -eq $printer.DriverName }

                        [PSCustomObject]@{
                            PrinterName  = $printer.Name
                            DriverName   = $printer.DriverName
                            PortName     = $printer.PortName
                            Published    = $printer.Published
                            Shared       = $printer.Shared
                            ShareName    = $printer.ShareName
                            Type         = $printer.Type
                            DeviceType   = $printer.DeviceType
                            Manufacturer = $driver.Manufacturer
                            InfPath      = $driver.InfPath
                        }
                    }

                    $combined
                }

                # Step 3 -- Execute locally or remotely
                if ($computer -eq $env:COMPUTERNAME) {
                    $printerInfo = & $scriptBlock
                } else {
                    $params = @{
                        ComputerName = $computer
                        ScriptBlock  = $scriptBlock
                    }
                    if ($Credential) { $params.Credential = $Credential }
                    $printerInfo = Invoke-Command @params
                }

                # Step 4 -- Emit results with metadata
                if ($printerInfo) {
                    foreach ($printer in $printerInfo) {
                        [PSCustomObject]@{
                            ComputerName   = $computer
                            PrinterName    = $printer.PrinterName
                            DriverName     = $printer.DriverName
                            PortName       = $printer.PortName
                            Published      = $printer.Published
                            Shared         = $printer.Shared
                            ShareName      = $printer.ShareName
                            Type           = $printer.Type
                            DeviceType     = $printer.DeviceType
                            Manufacturer   = $printer.Manufacturer
                            InfPath        = $printer.InfPath
                            Status         = 'Success'
                            CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                        }
                    }
                } else {
                    [PSCustomObject]@{
                        ComputerName   = $computer
                        PrinterName    = 'None'
                        Status         = 'No printers found'
                        CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                    }
                }
            }
            catch {
                [PSCustomObject]@{
                    ComputerName   = $computer
                    Error          = $_.Exception.Message
                    Status         = 'Failed'
                    CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
                }
            }
        }
    }
}
