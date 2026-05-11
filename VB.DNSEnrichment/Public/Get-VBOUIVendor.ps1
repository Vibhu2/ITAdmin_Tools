function Get-VBOUIVendor {
<#
.SYNOPSIS
    Look up the vendor organisation for a MAC address from the IEEE OUI database (Layer 11).

.DESCRIPTION
    Downloads the IEEE OUI CSV from https://standards-oui.ieee.org/oui/oui.csv on the
    first call if the file is missing, and refreshes it if the file is older than 30 days.
    After download, the CSV is imported into a script-scope hashtable keyed by the 6-char
    uppercase OUI prefix. All subsequent lookups within the session hit the hashtable --
    no file I/O per call.

    The vendor organisation name is also mapped to a device class hint via an ordered
    lookup table embedded in this function. The hints feed Resolve-VBDeviceClass Tier 13
    as a low-confidence fallback.

    Download uses TLS 1.2 (forced on PS 5.1). The file is saved with UTF-8 BOM encoding.

    Prerequisite: $MACAddress must not be null/empty. No network prerequisite beyond
    the initial download (which happens silently).

.PARAMETER MACAddress
    The MAC address to look up. Accepts any common separator format
    (00:1A:2B:3C:4D:5E, 00-1A-2B-3C-4D-5E, 001A.2B3C.4D5E, 001A2B3C4D5E).

.PARAMETER Context
    Environment context from Get-VBEnrichmentContext. Provides OUIFilePath.
    If omitted, the default path $env:LOCALAPPDATA\VB.DNSEnrichment\oui.csv is used.

.OUTPUTS
    [PSCustomObject] -- base layer result fields (IPAddress will be $null when called
    standalone without an IP) plus:
        Vendor           [string]   Raw IEEE organisation name
        VendorDeviceClass [string]  Mapped class hint for Resolve-VBDeviceClass
        MACNormalised    [string]   6-char uppercase OUI used for the lookup

.EXAMPLE
    $ctx = Get-VBEnrichmentContext
    Get-VBOUIVendor -MACAddress '00:1A:2B:3C:4D:5E' -Context $ctx

.EXAMPLE
    '00:1A:2B:3C:4D:5E' | Get-VBOUIVendor -Context $ctx

.NOTES
    Version:      1.0.0
    MinPSVersion: 5.1
    Author:       VB
    ChangeLog:
        1.0.0 -- 2026-05-11 -- Initial release
#>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$MACAddress,

        [Parameter()]
        [PSCustomObject]$Context,

        # IPAddress is optional -- passed by orchestrator so result object has context
        [Parameter()]
        [string]$IPAddress = $null
    )

    begin {
        $LAYER_NUM  = 11
        $LAYER_NAME = 'OUI'

        $ouiFilePath = if ($Context -and $Context.OUIFilePath) {
            $Context.OUIFilePath
        }
        else {
            Join-Path $env:LOCALAPPDATA 'VB.DNSEnrichment\oui.csv'
        }

        # Vendor -> DeviceClass hint table (ordered -- first match wins)
        $vendorClassMap = [ordered]@{
            'Yealink'          = 'IPPhone'
            'Poly'             = 'IPPhone'
            'Polycom'          = 'IPPhone'
            'Grandstream'      = 'IPPhone'
            'Snom'             = 'IPPhone'
            'Cisco Systems'    = 'NetworkDevice'
            'Ubiquiti'         = 'NetworkDevice'
            'Aruba'            = 'NetworkDevice'
            'Juniper'          = 'NetworkDevice'
            'Fortinet'         = 'NetworkDevice'
            'Hikvision'        = 'Camera'
            'Dahua'            = 'Camera'
            'Axis'             = 'Camera'
            'Hanwha'           = 'Camera'
            'Hewlett Packard'  = 'Workstation'
            'HP Inc'           = 'Workstation'
            'Dell'             = 'Workstation'
            'Apple'            = 'Workstation'
            'Lenovo'           = 'Workstation'
            'APC'              = 'UPS'
            'Eaton'            = 'UPS'
            'Synology'         = 'NAS'
            'QNAP'             = 'NAS'
        }

        # One-shot cache: load OUI table once per session
        if ($null -eq $Script:VBOUITable) {
            $Script:VBOUITable = Invoke-VBLoadOUITable -OUIFilePath $ouiFilePath
        }
    }

    process {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()

        if ([string]::IsNullOrWhiteSpace($MACAddress)) {
            $sw.Stop()
            return New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'Skipped' -ExecutionMs $sw.ElapsedMilliseconds `
                -SkipReason 'NoMACAvailable' `
                -Impact 'No vendor lookup -- MAC address was not obtained from any prior layer'
        }

        try {
            $macNorm = ConvertTo-VBNormalisedMAC -MACAddress $MACAddress
            if ($null -eq $macNorm) {
                $sw.Stop()
                return New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                    -Status 'Failed' -ExecutionMs $sw.ElapsedMilliseconds `
                    -ErrorDetail "MAC '$MACAddress' could not be normalised -- unexpected format"
            }

            # OUI is the first 6 chars of the 12-char normalised MAC
            $oui    = $macNorm.Substring(0, 6)
            $vendor = $Script:VBOUITable[$oui]

            if ([string]::IsNullOrWhiteSpace($vendor)) {
                $sw.Stop()
                return New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                    -Status 'NoResult' -ExecutionMs $sw.ElapsedMilliseconds `
                    -ExtraFields @{
                        Vendor            = $null
                        VendorDeviceClass = 'Unknown'
                        MACNormalised     = $oui
                    }
            }

            # Map vendor name to device class hint
            $vendorClass = 'Unknown'
            foreach ($key in $vendorClassMap.Keys) {
                if ($vendor -match [regex]::Escape($key)) {
                    $vendorClass = $vendorClassMap[$key]
                    break
                }
            }

            $sw.Stop()
            New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'Success' -ExecutionMs $sw.ElapsedMilliseconds `
                -ExtraFields @{
                    Vendor            = $vendor
                    VendorDeviceClass = $vendorClass
                    MACNormalised     = $oui
                }
        }
        catch {
            $sw.Stop()
            New-VBLayerResult -IPAddress $IPAddress -Layer $LAYER_NUM -LayerName $LAYER_NAME `
                -Status 'Failed' -ExecutionMs $sw.ElapsedMilliseconds `
                -ErrorDetail $_.Exception.Message
        }
    }
}

function Invoke-VBLoadOUITable {
<#
.SYNOPSIS
    Load (and if necessary download/refresh) the IEEE OUI CSV into a hashtable.
    Private helper used only by Get-VBOUIVendor.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$OUIFilePath
    )

    $OUI_URL       = 'https://standards-oui.ieee.org/oui/oui.csv'
    $REFRESH_DAYS  = 30
    $MIN_SIZE_BYTES = 1MB

    # Ensure parent folder exists
    $parent = Split-Path -Path $OUIFilePath -Parent
    if (-not (Test-Path -LiteralPath $parent)) {
        New-Item -Path $parent -ItemType Directory -Force | Out-Null
    }

    $needsDownload = $false

    if (-not (Test-Path -LiteralPath $OUIFilePath)) {
        Write-Verbose "[OUI] oui.csv not found -- downloading from IEEE"
        $needsDownload = $true
    }
    else {
        $item = Get-Item -LiteralPath $OUIFilePath
        if ($item.Length -lt $MIN_SIZE_BYTES) {
            Write-Verbose "[OUI] oui.csv too small ($([int]($item.Length/1KB)) KB) -- re-downloading"
            $needsDownload = $true
        }
        elseif (((Get-Date) - $item.CreationTime).TotalDays -gt $REFRESH_DAYS) {
            Write-Verbose "[OUI] oui.csv is $([int]((Get-Date) - $item.CreationTime).TotalDays) days old (threshold $REFRESH_DAYS) -- refreshing"
            $needsDownload = $true
        }
    }

    if ($needsDownload) {
        try {
            # Force TLS 1.2 on PS 5.1
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

            Write-Verbose "[OUI] Downloading OUI database from $OUI_URL ..."
            $tmpPath = "$OUIFilePath.tmp"
            Invoke-WebRequest -Uri $OUI_URL -OutFile $tmpPath -UseBasicParsing -ErrorAction Stop

            # Validate the download before replacing
            $tmpItem = Get-Item -LiteralPath $tmpPath
            if ($tmpItem.Length -lt $MIN_SIZE_BYTES) {
                Remove-Item -LiteralPath $tmpPath -Force -ErrorAction SilentlyContinue
                throw "Downloaded file too small ($([int]($tmpItem.Length/1KB)) KB) -- download may have failed"
            }

            # Replace old file
            if (Test-Path -LiteralPath $OUIFilePath) {
                Remove-Item -LiteralPath $OUIFilePath -Force
            }
            Move-Item -LiteralPath $tmpPath -Destination $OUIFilePath -Force
            Write-Verbose "[OUI] Download complete: $([math]::Round($tmpItem.Length/1MB,1)) MB"
        }
        catch {
            Write-Warning "[OUI] Download failed: $($_.Exception.Message). Vendor lookup will be unavailable this session."
            return @{}
        }
    }

    # Import CSV into hashtable keyed by 6-char uppercase OUI (Assignment column)
    Write-Verbose "[OUI] Loading OUI table into memory..."
    $table = @{}
    try {
        $rows = Import-Csv -LiteralPath $OUIFilePath -Encoding UTF8 -ErrorAction Stop
        foreach ($row in $rows) {
            # IEEE CSV columns: Registry, Assignment, "Organization Name", "Organization Address"
            $oui    = $row.Assignment.ToUpperInvariant().Trim()
            $vendor = $row.'Organization Name'.Trim()
            if ($oui.Length -eq 6 -and -not [string]::IsNullOrWhiteSpace($vendor)) {
                $table[$oui] = $vendor
            }
        }
        Write-Verbose "[OUI] Table loaded: $($table.Count) OUI entries"
    }
    catch {
        Write-Warning "[OUI] Failed to parse oui.csv: $($_.Exception.Message)"
    }

    $table
}
