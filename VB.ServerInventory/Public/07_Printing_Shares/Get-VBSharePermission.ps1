# ============================================================
# FUNCTION : Get-VBSmbSharePermission
# MODULE   : VBTools
# VERSION  : 1.0.1
# CHANGED  : 29-05-2026 -- Initial release
# AUTHOR   : Vibhu
# PURPOSE  : Returns SMB share-level and NTFS top-level ACEs as objects
# ENCODING : UTF-8 with BOM -- do not re-save without BOM
# ============================================================

function Get-VBSmbSharePermission {
<#
.SYNOPSIS
    Returns SMB share-level and NTFS top-level ACEs for non-administrative
    file shares on the local machine as structured objects.

.DESCRIPTION
    Queries Get-SmbShare and filters out administrative shares (ADMIN$, IPC$,
    drive-letter shares, SYSVOL, NETLOGON, PRINT$) and system paths under
    C:\Windows. For every qualifying share, two rows are emitted per ACE --
    one for share-level permissions and one for NTFS top-level permissions.

    Output is pipeline-friendly: pipe to Export-VBSmbSharePermissionReport
    for a grouped text report, or to Export-Csv for spreadsheet analysis.

    Errors on individual shares are captured per-row and do not stop
    processing of remaining shares.

.PARAMETER ShareName
    One or more share names to query. Accepts pipeline input by value and
    by property name. Defaults to all qualifying shares when omitted.

.EXAMPLE
    Get-VBSmbSharePermission

    Returns all share and NTFS ACEs for every non-administrative SMB share
    on the local machine.

.EXAMPLE
    Get-VBSmbSharePermission -ShareName 'Data', 'Finance'

    Returns ACEs only for the shares named Data and Finance.

.EXAMPLE
    Get-VBSmbSharePermission | Export-Csv -Path C:\Reports\SharePerms.csv -NoTypeInformation -Encoding UTF8

    Exports all ACEs to a CSV file for further analysis.

.EXAMPLE
    Get-VBSmbSharePermission | Where-Object { $_.PermissionLayer -eq 'NTFS' -and -not $_.IsInherited }

    Filters to explicit (non-inherited) NTFS ACEs only.

.EXAMPLE
    $OutputPath = Join-Path 'C:\Realtime' "SharePermissions_$(Get-Date -Format 'yyyyMMdd').txt"
    Get-VBSmbSharePermission | Export-VBSmbSharePermissionReport -OutputPath $OutputPath

    Produces a grouped, human-readable text report of all share permissions.

.OUTPUTS
    [PSCustomObject]

    Each object represents a single ACE with the following properties:

        ComputerName      [string]  -- Machine the share resides on
        Status            [string]  -- 'Success' or 'Failed'
        ShareName         [string]  -- SMB share name
        SharePath         [string]  -- Local file system path
        ShareDescription  [string]  -- Share description (may be empty)
        PermissionLayer   [string]  -- 'Share' or 'NTFS'
        Identity          [string]  -- Account or group name
        AccessControlType [string]  -- 'Allow' or 'Deny'
        Rights            [string]  -- Permission level or FileSystemRights value
        IsInherited       [bool]    -- True if ACE is inherited; $null for Share-layer rows
        Error             [string]  -- $null on success; exception message on failure
        CollectionTime    [string]  -- Timestamp in dd-MM-yyyy HH:mm:ss format

.NOTES
    Version  : 1.0.0
    Author   : Vibhu
    Modified : 29-05-2026
    Category : Windows Server Administration

    This function is local-only and does not use WinRM or remoting.
    Run directly on the target machine.

    Requires the SmbShare module (available by default on Windows Server 2012+
    and Windows 8+).
#>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string[]]$ShareName
    )

    begin {
        $CollectedNames = [System.Collections.Generic.List[string]]::new()
    }

    process {
        if ($ShareName) {
            foreach ($name in $ShareName) { $CollectedNames.Add($name) }
        }
    }

    end {
        # Step 1 -- Resolve qualifying shares
        Write-Verbose '[Get-VBSmbSharePermission] Querying SMB shares'

        $allShares = Get-SmbShare | Where-Object {
            $_.ShareType -eq 'FileSystemDirectory' -and
            $_.Path -and
            $_.Path -notmatch '^C:\\($|Windows)' -and
            $_.Name -notmatch '^(ADMIN\$|IPC\$|PRINT\$|SYSVOL|NETLOGON|[A-Z]\$)$'
        }

        if ($CollectedNames.Count -gt 0) {
            $allShares = $allShares | Where-Object { $CollectedNames -contains $_.Name }
        }

        if ($null -eq $allShares -or @($allShares).Count -eq 0) {
            Write-Warning '[Get-VBSmbSharePermission] No qualifying shares found'
            return
        }

        $collectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')

        # Step 2 -- Emit one object per ACE for each share
        foreach ($share in $allShares) {
            Write-Verbose "[Get-VBSmbSharePermission] Processing share: $($share.Name)"

            # Share-level ACEs
            try {
                $smbAces = Get-SmbShareAccess -Name $share.Name
                if ($null -eq $smbAces -or @($smbAces).Count -eq 0) {
                    Write-Warning "[Get-VBSmbSharePermission] No share-level ACEs found for: $($share.Name)"
                } else {
                    foreach ($ace in $smbAces) {
                        [PSCustomObject]@{
                            ComputerName      = $env:COMPUTERNAME
                            Status            = 'Success'
                            ShareName         = $share.Name
                            SharePath         = $share.Path
                            ShareDescription  = $share.Description
                            PermissionLayer   = 'Share'
                            Identity          = $ace.AccountName
                            AccessControlType = $ace.AccessControlType.ToString()
                            Rights            = $ace.AccessRight.ToString()
                            IsInherited       = $null
                            Error             = $null
                            CollectionTime    = $collectionTime
                        }
                    }
                }
            }
            catch {
                Write-Warning "[Get-VBSmbSharePermission] Share-level ACL query failed for: $($share.Name)"
                [PSCustomObject]@{
                    ComputerName      = $env:COMPUTERNAME
                    Status            = 'Failed'
                    ShareName         = $share.Name
                    SharePath         = $share.Path
                    ShareDescription  = $share.Description
                    PermissionLayer   = 'Share'
                    Identity          = $null
                    AccessControlType = $null
                    Rights            = $null
                    IsInherited       = $null
                    Error             = $_.Exception.Message
                    CollectionTime    = $collectionTime
                }
            }

            # NTFS top-level ACEs
            try {
                $ntfsAces = (Get-Acl -Path $share.Path).Access
                if ($null -eq $ntfsAces -or @($ntfsAces).Count -eq 0) {
                    Write-Warning "[Get-VBSmbSharePermission] No NTFS ACEs found at path: $($share.Path)"
                } else {
                    foreach ($ace in $ntfsAces) {
                        [PSCustomObject]@{
                            ComputerName      = $env:COMPUTERNAME
                            Status            = 'Success'
                            ShareName         = $share.Name
                            SharePath         = $share.Path
                            ShareDescription  = $share.Description
                            PermissionLayer   = 'NTFS'
                            Identity          = $ace.IdentityReference.ToString()
                            AccessControlType = $ace.AccessControlType.ToString()
                            Rights            = $ace.FileSystemRights.ToString()
                            IsInherited       = $ace.IsInherited
                            Error             = $null
                            CollectionTime    = $collectionTime
                        }
                    }
                }
            }
            catch {
                Write-Warning "[Get-VBSmbSharePermission] NTFS ACL query failed for path: $($share.Path)"
                [PSCustomObject]@{
                    ComputerName      = $env:COMPUTERNAME
                    Status            = 'Failed'
                    ShareName         = $share.Name
                    SharePath         = $share.Path
                    ShareDescription  = $share.Description
                    PermissionLayer   = 'NTFS'
                    Identity          = $null
                    AccessControlType = $null
                    Rights            = $null
                    IsInherited       = $null
                    Error             = $_.Exception.Message
                    CollectionTime    = $collectionTime
                }
            }
        }
    }
}


# ============================================================
# FUNCTION : Export-VBSmbSharePermissionReport
# MODULE   : VBTools
# VERSION  : 1.0.0
# CHANGED  : 29-05-2026 -- Initial release
# AUTHOR   : Vibhu
# PURPOSE  : Writes Get-VBSmbSharePermission output to a grouped text report
# ENCODING : UTF-8 with BOM -- do not re-save without BOM
# ============================================================

function Export-VBSmbSharePermissionReport {
<#
.SYNOPSIS
    Writes output from Get-VBSmbSharePermission to a grouped, human-readable
    text report file.

.DESCRIPTION
    Consumes PSCustomObject rows from Get-VBSmbSharePermission and produces a
    text file grouped by share. Each share section shows share-level ACEs
    followed by NTFS top-level ACEs. Errors are surfaced inline under the
    relevant layer heading.

    This function buffers all pipeline input before writing -- required because
    grouping by share name is intrinsically blocking.

.PARAMETER Data
    Output from Get-VBSmbSharePermission. Accepts pipeline input.

.PARAMETER OutputPath
    Full path for the output text file. Parent directory must already exist.

.EXAMPLE
    Get-VBSmbSharePermission | Export-VBSmbSharePermissionReport -OutputPath 'C:\Reports\SharePerms.txt'

    Runs the query and writes a grouped report to the specified path.

.EXAMPLE
    $OutputPath = Join-Path 'C:\Realtime' "DSI-DH01-DC-003_SharePermissions_$(Get-Date -Format 'yyyyMMdd').txt"
    Get-VBSmbSharePermission | Export-VBSmbSharePermissionReport -OutputPath $OutputPath
    Write-Host "Saved to: $OutputPath"

    Writes the report with a date-stamped filename matching the original script convention.

.EXAMPLE
    $data = Get-VBSmbSharePermission
    $data | Export-VBSmbSharePermissionReport -OutputPath 'C:\Reports\SharePerms.txt'
    $data | Export-Csv -Path 'C:\Reports\SharePerms.csv' -NoTypeInformation -Encoding UTF8

    Reuse the collected objects for both a readable report and a CSV export.

.OUTPUTS
    None. This function writes a file and does not return objects to the pipeline.

.NOTES
    Version  : 1.0.0
    Author   : Vibhu
    Modified : 29-05-2026
    Category : Windows Server Administration

    Output file is written as UTF-8. Parent directory must exist before calling.
    This function intentionally buffers all input -- grouping requires it.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [PSCustomObject[]]$Data,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    begin {
        $all = [System.Collections.Generic.List[PSCustomObject]]::new()
    }

    process {
        foreach ($row in $Data) { $all.Add($row) }
    }

    end {
        if ($all.Count -eq 0) {
            Write-Warning '[Export-VBSmbSharePermissionReport] No data received -- report not written'
            return
        }

        Write-Verbose "[Export-VBSmbSharePermissionReport] Building report for $($all.Count) ACE rows"

        $lines  = [System.Collections.Generic.List[string]]::new()
        $server = ($all | Select-Object -First 1).ComputerName
        $ts     = ($all | Select-Object -First 1).CollectionTime

        $lines.Add("Share Permission Report")
        $lines.Add("Server         : $server")
        $lines.Add("Generated      : $ts")
        $lines.Add("Total ACE rows : $($all.Count)")
        $lines.Add("=" * 60)

        $shareNames = $all | Select-Object -ExpandProperty ShareName -Unique

        foreach ($shareName in $shareNames) {
            $shareRows = $all | Where-Object { $_.ShareName -eq $shareName }
            $first     = $shareRows | Select-Object -First 1

            $lines.Add("")
            $lines.Add("=== Share: $shareName ===")
            $lines.Add("Path        : $($first.SharePath)")
            $lines.Add("Description : $($first.ShareDescription)")

            # Share-level ACEs
            $lines.Add("")
            $lines.Add("-- Share-Level Permissions --")
            $shareAces = $shareRows | Where-Object { $_.PermissionLayer -eq 'Share' }

            if (@($shareAces).Count -eq 0) {
                $lines.Add("  (none)")
            } else {
                foreach ($ace in $shareAces) {
                    if ($ace.Error) {
                        $lines.Add("  [ERROR] $($ace.Error)")
                    } else {
                        $lines.Add("  $($ace.Identity)  |  $($ace.AccessControlType)  |  $($ace.Rights)")
                    }
                }
            }

            # NTFS ACEs
            $lines.Add("")
            $lines.Add("-- NTFS Permissions (top level) --")
            $ntfsAces = $shareRows | Where-Object { $_.PermissionLayer -eq 'NTFS' }

            if (@($ntfsAces).Count -eq 0) {
                $lines.Add("  (none)")
            } else {
                foreach ($ace in $ntfsAces) {
                    if ($ace.Error) {
                        $lines.Add("  [ERROR] $($ace.Error)")
                    } else {
                        $inherited = if ($ace.IsInherited) { 'Inherited' } else { 'Explicit' }
                        $lines.Add("  $($ace.Identity)  |  $($ace.AccessControlType)  |  $($ace.Rights)  |  $inherited")
                    }
                }
            }
        }

        $lines | Set-Content -Path $OutputPath -Encoding UTF8
        Write-Verbose "[Export-VBSmbSharePermissionReport] Report written to: $OutputPath"
    }
}
