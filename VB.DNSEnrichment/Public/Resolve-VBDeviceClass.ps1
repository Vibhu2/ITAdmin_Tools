function Resolve-VBDeviceClass {
<#
.SYNOPSIS
    Classify an enriched IP address into a device class from collected signal data.

.DESCRIPTION
    Pure classification logic -- never calls any probe function. Takes the merged
    signals from all layer outputs (OSClass, open ports, HTTP/SNMP/RTSP/mDNS strings,
    OUI vendor) and returns a DeviceClass, Confidence, and DeviceClassSource.

    Classification runs through a tiered switch in priority order:
        Tier 1  -- AD OSClass (authoritative -- domain-joined machines)
        Tier 2  -- Cameras   (RTSP port 554, banner, or vendor string)
        Tier 3  -- VoIP      (SIP ports 5060/5061, or vendor string)
        Tier 4  -- Printers  (JetDirect 9100, IPP 631, or vendor string)
        Tier 5  -- Scanners  (mDNS _scanner service type)
        Tier 6  -- VMware    (ESXi ports 902/9443, or banner)
        Tier 7  -- NAS       (vendor/banner string match)
        Tier 8  -- UPS/PDU   (vendor/banner string match)
        Tier 9  -- Network   (Cisco/Ubiquiti/Aruba etc.)
        Tier 10 -- iOS       (port 62078)
        Tier 11 -- IoT/MQTT  (port 1883)
        Tier 12 -- Windows fallback (RDP+RPC = Workstation, WinRM+SMB = Server)
        Tier 13 -- OUI-only  (lowest-signal fallback)
        default -- Unknown

    This function does NOT enforce the skip-if-resolved gate -- that is the
    orchestrator's responsibility.

.PARAMETER OSClass
    From Get-VBADComputer. 'DomainController', 'Server', or 'Workstation'.

.PARAMETER OpenPorts
    Comma-separated port list from Get-VBTCPFingerprint (e.g. '80,443,9100').

.PARAMETER HTTPTitle
    Page title from Get-VBHTTPBanner.

.PARAMETER HTTPServer
    Server header from Get-VBHTTPBanner.

.PARAMETER SNMPDescr
    sysDescr OID value from Get-VBSNMPIdentity.

.PARAMETER RTSPBanner
    Banner from Get-VBRTSPBanner.

.PARAMETER OUIVendor
    Organization name from Get-VBOUIVendor.

.PARAMETER VendorDeviceClass
    Pre-mapped class hint from Get-VBOUIVendor lookup table.

.PARAMETER MDNSServiceType
    Service type string from Get-VBmDNSRecord (e.g. '_scanner._tcp').

.OUTPUTS
    [PSCustomObject]
        DeviceClass       [string]
        Confidence        [string]   High | Medium | Low | None
        DeviceClassSource [string]   comma-separated list of signals that matched

.EXAMPLE
    Resolve-VBDeviceClass -OSClass 'Server' -OpenPorts '135,445,3389,5985'

.EXAMPLE
    Resolve-VBDeviceClass -OpenPorts '9100,631' -OUIVendor 'HP Inc.'

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
        [Parameter()]
        [string]$OSClass,

        [Parameter()]
        [string]$OpenPorts,

        [Parameter()]
        [string]$HTTPTitle,

        [Parameter()]
        [string]$HTTPServer,

        [Parameter()]
        [string]$SNMPDescr,

        [Parameter()]
        [string]$RTSPBanner,

        [Parameter()]
        [string]$OUIVendor,

        [Parameter()]
        [string]$VendorDeviceClass,

        [Parameter()]
        [string]$MDNSServiceType
    )

    # Combine free-text signals into one lowercase string for regex matching
    $combined = (
        "$HTTPTitle $HTTPServer $SNMPDescr $RTSPBanner $OUIVendor"
    ).ToLower().Trim()

    # Parse open ports into an int array
    $ports = @()
    if (-not [string]::IsNullOrWhiteSpace($OpenPorts)) {
        $ports = @(
            $OpenPorts -split ',' |
                ForEach-Object {
                    $t = $_.Trim()
                    if ($t -match '^\d+$') { [int]$t }
                }
        )
    }

    $sources = New-Object System.Collections.Generic.List[string]

    $class = switch ($true) {

        # Tier 1 -- AD OSClass is authoritative
        ($OSClass -eq 'DomainController') {
            $sources.Add('OSClass')
            'DomainController'
            break
        }
        ($OSClass -eq 'Server') {
            $sources.Add('OSClass')
            'Server'
            break
        }
        ($OSClass -eq 'Workstation') {
            $sources.Add('OSClass')
            'Workstation'
            break
        }

        # Tier 2 -- Cameras (RTSP is exclusive to cameras/NVR)
        (
            $ports -contains 554 -or
            $RTSPBanner -match 'RTSP' -or
            $combined -match 'hikvision|dahua|axis|hanwha|bosch.*security|milestone|reolink|amcrest'
        ) {
            $sources.Add('RTSPOrVendor')
            'Camera'
            break
        }

        # Tier 3 -- VoIP / IP phones
        (
            $ports -contains 5060 -or
            $ports -contains 5061 -or
            $combined -match 'yealink|polycom|grandstream|snom|sip.*phone|voip|cisco.*phone|cisco ip'
        ) {
            $sources.Add('SIPOrVendor')
            'IPPhone'
            break
        }

        # Tier 4 -- Printers
        (
            $ports -contains 9100 -or
            $ports -contains 631 -or
            $ports -contains 515 -or
            $MDNSServiceType -match '_ipp\._tcp|_pdl-datastream\._tcp|_ipps\._tcp' -or
            $combined -match 'jetdirect|kyocera|ricoh|xerox|laserjet|bizhub|taskalfa|workcentre|brother|lexmark|konica|sharp.*mx|develop.*ineo'
        ) {
            $sources.Add('PrinterPortOrVendor')
            'Printer'
            break
        }

        # Tier 5 -- Scanners (mDNS service type is the primary signal)
        (
            $MDNSServiceType -match '_scanner' -or
            $combined -match '\bscanner\b'
        ) {
            $sources.Add('mDNSScanner')
            'Scanner'
            break
        }

        # Tier 6 -- VMware ESXi / virtual hosts
        (
            $ports -contains 902 -or
            $ports -contains 9443 -or
            $combined -match 'esxi|vmware|vsphere|vcenter'
        ) {
            $sources.Add('ESXiPortOrBanner')
            'VirtualHost'
            break
        }

        # Tier 7 -- NAS
        (
            $combined -match 'synology|qnap|nas|diskstation|truenas|freenas|readynas|drobo'
        ) {
            $sources.Add('NASBanner')
            'NAS'
            break
        }

        # Tier 8 -- UPS / PDU
        (
            $combined -match '\bapc\b|eaton|\bups\b|\bpdu\b|powerware|tripplite|cyberpower'
        ) {
            $sources.Add('UPSBanner')
            'UPS'
            break
        }

        # Tier 9 -- Network infrastructure
        (
            $combined -match 'cisco|ubiquiti|aruba|juniper|fortinet|mikrotik|unifi|palo.*alto|sonicwall|meraki|netgear|draytek' -or
            $combined -match '\bswitch\b|\brouter\b|\bfirewall\b|\baccess.point\b|\bwap\b'
        ) {
            $sources.Add('NetworkVendorOrBanner')
            'NetworkDevice'
            break
        }

        # Tier 10 -- iOS (iTunes sync port)
        ($ports -contains 62078) {
            $sources.Add('iOSPort')
            'Mobile'
            break
        }

        # Tier 11 -- IoT / MQTT broker
        ($ports -contains 1883) {
            $sources.Add('MQTTPort')
            'IoT'
            break
        }

        # Tier 12 -- Windows fingerprint fallback (no AD)
        ($ports -contains 3389 -and $ports -contains 135) {
            $sources.Add('WindowsPorts')
            'Workstation'
            break
        }
        ($ports -contains 5985 -and $ports -contains 445) {
            $sources.Add('WindowsPorts')
            'Server'
            break
        }

        # Tier 13 -- OUI vendor hint (lowest signal -- no ports, no banners)
        (
            -not [string]::IsNullOrWhiteSpace($VendorDeviceClass) -and
            $VendorDeviceClass -ne 'Unknown'
        ) {
            $sources.Add('OUIVendor')
            $VendorDeviceClass
            break
        }

        default {
            'Unknown'
        }
    }

    $confidence = switch ($class) {
        'DomainController' { 'High' }
        'Server'           { 'High' }
        'Workstation'      { 'High' }
        'IPPhone'          { 'High' }
        'Printer'          { 'High' }
        'Camera'           {
            if ($ports -contains 554) { 'High' } else { 'Medium' }
        }
        'Unknown'          { 'None' }
        default            { 'Medium' }
    }

    [PSCustomObject]@{
        DeviceClass       = $class
        Confidence        = $confidence
        DeviceClassSource = ($sources -join ',')
    }
}
