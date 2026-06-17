# ============================================================
# FUNCTION : Install-VBPrinterDriver
# MODULE   : VB.WorkstationReport
# VERSION  : 1.2.0
# CHANGED  : 16-06-2026 -- Standards fix: output property order, [OutputType], CollectionTime,
#                          Error property; downgrade post-verify throw to Write-Warning on reboot-required
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Stages and installs a printer driver into the local driver store
# ENCODING : UTF-8 with BOM
# NOTE     : Private -- called by Add-VBUserPrinter and Set-VBUserPrinterMigration
# ============================================================
# CHANGELOG (newest first):
# v1.2.0 -- 16-06-2026 -- Output property order fixed (ComputerName, Status, Error, CollectionTime);
#                         [OutputType] added; post-verify downgraded to Write-Warning when RebootRequired.
# v1.1.0 -- 16-06-2026 -- Post-install verify; 3010 handling; WhatIf fix; retry; verbose output;
#                         Duration + RebootRequired; extrac32 cab fallback.
# v1.0.0 -- 16-06-2026 -- Initial release (extracted from Add-VBUserPrinter v1.2.0).
# ============================================================

function Install-VBPrinterDriver {
    <#
    .SYNOPSIS
        Stages and installs a printer driver into the local driver store.

    .DESCRIPTION
        Install-VBPrinterDriver checks whether a named printer driver is already present
        in the local driver store. If found it returns immediately with Action = 'AlreadyInstalled'.

        If the driver is not installed and -DriverPath is supplied, it stages the driver
        using pnputil then registers it with Add-PrinterDriver. Three source formats are
        accepted:

          .inf file   -- pnputil stages the single .inf directly.
          .cab file   -- expanded to a temp folder; pnputil /subdirs stages all .inf files inside.
                         expand.exe is tried first; extrac32.exe is used as a fallback for
                         vendor .cab files (HP especially) that expand.exe cannot handle.
          Folder      -- pnputil /subdirs recurses the entire folder tree.

        The pnputil staging call and the Add-PrinterDriver registration are each wrapped in a
        retry loop (max 3 attempts, 5-second pause between attempts) so transient driver-store
        locks held by Windows Update or another installer do not cause a hard failure.

        pnputil exit code 3010 ("staged successfully, reboot required") is treated as success.
        The function sets RebootRequired = $true and emits a warning. The post-install driver
        store verification is downgraded to Write-Warning (not a throw) in the reboot-required
        case, because the print subsystem may not reflect the staged driver until after restart.

        NOTE: .exe and .msi installers are NOT supported -- extract them first (e.g.
        setup.exe /extract C:\Temp\Driver or 7-Zip), then pass the extracted folder.

        This is a private helper. Use Add-VBUserPrinter or Set-VBUserPrinterMigration instead.

    .PARAMETER DriverName
        Exact driver name as it appears in the driver store.
        Run: Get-PrinterDriver | Select-Object Name  to list installed drivers.

    .PARAMETER DriverPath
        Path to the driver source. Accepts a .inf file, .cab file, or extracted folder.
        Required when the driver is not already installed. Omit when the driver is
        already present -- the function skips installation silently.

    .OUTPUTS
        [PSCustomObject] with properties (in order):
          ComputerName   [string]  -- Always $env:COMPUTERNAME (driver install is local-only).
          Status         [string]  -- 'Success'. Throws on failure -- no 'Failed' object emitted.
          DriverName     [string]  -- The driver name supplied.
          DriverPath     [string]  -- The driver source path supplied.
          Action         [string]  -- 'AlreadyInstalled', 'Installed', or 'WhatIf'.
          RebootRequired [bool]    -- $true when pnputil returned 3010 (restart needed).
          Duration       [string]  -- Elapsed install time formatted hh:mm:ss.
          Error          [string]  -- Always $null on success (present for schema consistency).
          CollectionTime [string]  -- Timestamp dd-MM-yyyy HH:mm:ss.
        Throws a terminating error on failure.

    .NOTES
        Version  : 1.2.0
        Author   : Vibhu Bhatnagar
        Category : Printer Management (Private)
        Requires : Administrative privileges; PrintManagement module
    #>

    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory)]
        [string]$DriverName,

        [string]$DriverPath
    )

    if (Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue) {
        Write-Verbose "'$DriverName': already installed -- skipping."
        return [PSCustomObject]@{
            ComputerName   = $env:COMPUTERNAME
            Status         = 'Success'
            DriverName     = $DriverName
            DriverPath     = $DriverPath
            Action         = 'AlreadyInstalled'
            RebootRequired = $false
            Duration       = ([TimeSpan]::Zero).ToString('hh\:mm\:ss')
            Error          = $null
            CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
        }
    }

    if (-not $DriverPath) {
        throw "Driver '$DriverName' is not installed and no -DriverPath was supplied. " +
              "Supply -DriverPath as a .inf file, .cab file, or extracted folder. " +
              "Run: Get-PrinterDriver | Select-Object Name  to see what is installed."
    }

    if (-not (Test-Path -Path $DriverPath)) {
        throw "DriverPath not found: '$DriverPath'"
    }

    $startTime      = Get-Date
    $rebootRequired = $false
    $didProcess     = $false

    Write-Verbose "'$DriverName': not installed -- staging from '$DriverPath'."

    $extension    = [System.IO.Path]::GetExtension($DriverPath).ToLower()
    $tempExpanded = $null

    try {
        # --- Step 1: Resolve pnputil arguments based on DriverPath format ---
        if ($extension -eq '.cab') {
            $tempExpanded = Join-Path $env:TEMP "VBDriverStage_$([System.IO.Path]::GetFileNameWithoutExtension($DriverPath))"
            New-Item -ItemType Directory -Path $tempExpanded -Force | Out-Null

            Write-Verbose "'$DriverName': expanding .cab with expand.exe to '$tempExpanded'."
            $expandOutput = & expand.exe -F:* $DriverPath $tempExpanded 2>&1
            if ($LASTEXITCODE -ne 0) {
                Write-Warning "'$DriverName': expand.exe failed (exit $LASTEXITCODE) -- trying extrac32.exe fallback."
                $extracOutput = & extrac32.exe /Y /E "$DriverPath" /L "$tempExpanded" 2>&1
                if ($LASTEXITCODE -ne 0) {
                    throw "Failed to extract '$DriverPath'. expand.exe: $expandOutput | extrac32.exe: $extracOutput"
                }
            }
            $pnpArgs = @('/add-driver', "$tempExpanded\*.inf", '/subdirs', '/install')
        }
        elseif ($extension -eq '.inf') {
            $pnpArgs = @('/add-driver', $DriverPath, '/install')
        }
        elseif ((Get-Item -Path $DriverPath).PSIsContainer) {
            $pnpArgs = @('/add-driver', "$($DriverPath.TrimEnd('\'))\*.inf", '/subdirs', '/install')
        }
        else {
            throw "Unsupported DriverPath format '$extension'. Supply a .inf file, .cab file, or an extracted folder. " +
                  "For .exe or .msi installers, extract them first (e.g. setup.exe /extract C:\Temp\Driver) then pass the folder."
        }

        if ($PSCmdlet.ShouldProcess($DriverName, 'Stage and install printer driver')) {
            $didProcess = $true

            # --- Step 2: Stage with pnputil (retry on transient driver-store lock) ---
            $attempt = 0
            $staged  = $false
            while (-not $staged) {
                $attempt++
                try {
                    $pnpOutput = & pnputil @pnpArgs 2>&1
                    $pnpExit   = $LASTEXITCODE

                    foreach ($line in $pnpOutput) { Write-Verbose "'$DriverName': pnputil -- $line" }

                    if ($pnpExit -eq 3010) {
                        $rebootRequired = $true
                        Write-Warning "'$DriverName': staged successfully but a reboot is required (pnputil exit 3010)."
                    }
                    elseif ($pnpExit -ne 0) {
                        throw "pnputil failed (exit $pnpExit): $pnpOutput"
                    }
                    $staged = $true
                }
                catch {
                    if ($attempt -ge 3) { throw "pnputil failed after 3 attempts: $($_.Exception.Message)" }
                    Write-Warning "'$DriverName': attempt $attempt of 3 failed -- retrying in 5s. $($_.Exception.Message)"
                    Start-Sleep -Seconds 5
                }
            }

            # --- Step 3: Register with Add-PrinterDriver (retry on transient driver-store lock) ---
            $attempt    = 0
            $registered = $false
            while (-not $registered) {
                $attempt++
                try {
                    Add-PrinterDriver -Name $DriverName -ErrorAction Stop
                    $registered = $true
                }
                catch {
                    if ($attempt -ge 3) { throw "Add-PrinterDriver failed after 3 attempts: $($_.Exception.Message)" }
                    Write-Warning "'$DriverName': attempt $attempt of 3 failed -- retrying in 5s. $($_.Exception.Message)"
                    Start-Sleep -Seconds 5
                }
            }

            # --- Step 4: Post-install verification ---
            # When a reboot is required (3010), the driver may not be visible until restart.
            # Downgrade to Write-Warning rather than throwing -- the staging succeeded.
            if (-not (Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue)) {
                if ($rebootRequired) {
                    Write-Warning "'$DriverName': not yet visible in driver store -- a system restart is required."
                }
                else {
                    throw "Driver '$DriverName' staged successfully but is not visible in the driver store. A system restart may be required."
                }
            }

            Write-Verbose "'$DriverName': installed successfully."
        }

        return [PSCustomObject]@{
            ComputerName   = $env:COMPUTERNAME
            Status         = 'Success'
            DriverName     = $DriverName
            DriverPath     = $DriverPath
            Action         = if ($didProcess) { 'Installed' } else { 'WhatIf' }
            RebootRequired = $rebootRequired
            Duration       = ((Get-Date) - $startTime).ToString('hh\:mm\:ss')
            Error          = $null
            CollectionTime = (Get-Date).ToString('dd-MM-yyyy HH:mm:ss')
        }
    }
    finally {
        if ($null -ne $tempExpanded -and (Test-Path -Path $tempExpanded)) {
            Remove-Item -Path $tempExpanded -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}
