function Set-VBRedirectedFolderPermissions
{
    <#
    .SYNOPSIS
    Sets NTFS permissions on folders for redirected folder scenarios i.e user folders eg C:\Users.

    .DESCRIPTION
    This function configures NTFS permissions on specified folders to support Windows folder redirection. 
    It removes existing non-inherited permissions and applies the standard redirected folder permission set 
    including CREATOR OWNER full control, SYSTEM full control, Domain Admins full control, and limited 
    Everyone permissions for creating folders and reading attributes.

    .PARAMETER FolderPath
    Specifies one or more folder paths to configure permissions on. Accepts pipeline input and supports 
    multiple paths. Path validation is performed to ensure folders exist before processing.

    .EXAMPLE
    Set-VBRedirectedFolderPermissions -FolderPath "C:\RedirectedFolders\Desktop"
    
    Sets redirected folder permissions on a single Desktop folder path.

    .EXAMPLE
    "C:\RedirectedFolders\Desktop", "C:\RedirectedFolders\Documents" | Set-VBRedirectedFolderPermissions
    
    Processes multiple folder paths via pipeline input, configuring permissions on both Desktop and Documents folders.

    .EXAMPLE
    Get-ChildItem "C:\RedirectedFolders" -Directory | Set-VBRedirectedFolderPermissions -WhatIf
    
    Shows what permissions would be set on all subdirectories under RedirectedFolders without making changes.

    .OUTPUTS
    PSCustomObject
    Returns objects with FolderPath, Status, and Error (if applicable) properties for each processed folder.

    .NOTES
    Version: 1.0
    Author: System Administrator
    Category: Windows Server Administration
    
    Requires elevated permissions to modify NTFS ACLs. Supports -WhatIf and -Confirm parameters for safe testing.
    #>

    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('Path', 'FullName')]
        [string[]]$FolderPath
    )

    process
    {
        foreach ($path in $FolderPath)
        {
            if (-not (Test-Path $path))
            {
                [PSCustomObject]@{
                    FolderPath = $path
                    Status     = 'Failed'
                    Error      = 'Path does not exist'
                }
                continue
            }

            if ($PSCmdlet.ShouldProcess($path, "Set redirected folder permissions"))
            {
                try
                {
                    $acl = Get-Acl $path
                    
                    # Remove non-inherited permissions
                    $acl.Access | Where-Object { -not $_.IsInherited } | ForEach-Object { 
                        $acl.RemoveAccessRule($_) 
                    }

                    # Add required permissions
                    @(
                        @('CREATOR OWNER', 'FullControl', 'ContainerInherit,ObjectInherit'),
                        @('SYSTEM', 'FullControl', 'ContainerInherit,ObjectInherit'),
                        @('Domain Admins', 'FullControl', 'ContainerInherit,ObjectInherit'),
                        @('Everyone', 'CreateFolders,AppendData', 'None'),
                        @('Everyone', 'ListDirectory,ReadData', 'None'),
                        @('Everyone', 'ReadAttributes', 'None'),
                        @('Everyone', 'Traverse', 'None')
                    ) | ForEach-Object {
                        $rule = New-Object System.Security.AccessControl.FileSystemAccessRule($_[0], $_[1], $_[2], 'None', 'Allow')
                        $acl.AddAccessRule($rule)
                    }

                    Set-Acl -Path $path -AclObject $acl
                    
                    [PSCustomObject]@{
                        FolderPath = $path
                        Status     = 'Success'
                        PSTypeName = 'VB.RedirectedFolderPermissions'
                    }
                }
                catch
                {
                    [PSCustomObject]@{
                        FolderPath = $path
                        Status     = 'Failed'
                        Error      = $_.Exception.Message
                        PSTypeName = 'VB.RedirectedFolderPermissions'
                    }
                }
            }
        }
    }
}   

<# Original code for reference, not executed
function Set-RedirectedFolderPermissions {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string[]]$FolderPaths
    )

    begin {
        Write-Verbose "Initializing permission setting operation..."
    }

    process {
        foreach ($FolderPath in $FolderPaths) {
            if (-not (Test-Path $FolderPath)) {
                Write-Warning "Folder does not exist: $FolderPath"
                continue
            }

            if ($PSCmdlet.ShouldProcess($FolderPath, "Set NTFS permissions for redirected folder")) {
                try {
                    Write-Verbose "Getting ACL for $FolderPath"
                    $acl = Get-Acl $FolderPath

                    # Remove non-inherited rules
                    $nonInherited = $acl.Access | Where-Object { -not $_.IsInherited }
                    foreach ($entry in $nonInherited) {
                        $acl.RemoveAccessRule($entry)
                    }

                    # Create required access rules
                    $accessRules = @(
                        New-Object System.Security.AccessControl.FileSystemAccessRule("CREATOR OWNER", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow"),
                        New-Object System.Security.AccessControl.FileSystemAccessRule("SYSTEM", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow"),
                        New-Object System.Security.AccessControl.FileSystemAccessRule("Domain Admins", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow"),
                        New-Object System.Security.AccessControl.FileSystemAccessRule("Everyone", "CreateFolders,AppendData", "None", "None", "Allow"),
                        New-Object System.Security.AccessControl.FileSystemAccessRule("Everyone", "ListDirectory,ReadData", "None", "None", "Allow"),
                        New-Object System.Security.AccessControl.FileSystemAccessRule("Everyone", "ReadAttributes", "None", "None", "Allow"),
                        New-Object System.Security.AccessControl.FileSystemAccessRule("Everyone", "Traverse", "None", "None", "Allow")
                    )

                    foreach ($rule in $accessRules) {
                        $acl.AddAccessRule($rule)
                    }

                    Set-Acl -Path $FolderPath -AclObject $acl
                    Write-Host "✔ Permissions successfully set for: $FolderPath" -ForegroundColor Green
                }
                catch {
                    Write-Error "❌ Failed to set permissions on $FolderPath. Error: $_"
                }
            }
        }
    }

    end {
        Write-Verbose "Completed all folder operations."
    }
}
#>