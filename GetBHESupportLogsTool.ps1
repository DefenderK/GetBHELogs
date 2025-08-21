<#
.SYNOPSIS
    Collects logs and configuration relevant to BloodHound Enterprise troubleshooting.

.DESCRIPTION
    This tool exports Windows Application and System event logs, collects
    BloodHound Enterprise log files and settings from the profile of the account
    running the SHDelegator service, and compresses the results into a zip archive.

    It displays a simple text UI with progress for each item, a summary of collected
    and missing items, and options to open the output folder or the zip file.

    NOTE: Log level changes (LogLevel and EnumerationLogLevel) and service restarts
    are controlled only via command-line parameters; there are no interactive prompts
    for these settings.

.PARAMETER OutputRoot
    Root directory where the output folder and zip will be created. Defaults to the
    current user's Desktop.

.PARAMETER ServiceName
    The Windows service name to resolve the run-as account from. Defaults to
    'SHDelegator'. The script also attempts to match by DisplayName
    ('SharpHoundDelegator') or Description ('SharpHound Delegation Service').

.PARAMETER AutoStart
    Skip the initial "Press Enter / Help / Quit" screen and start immediately.

.PARAMETER ExcludeEventLogs
    Do not collect Windows Application and System event logs.

.PARAMETER ExcludeSettings
    Do not collect settings.json and skip any setting changes.

.PARAMETER SetLogLevel
    Sets the BHE settings.json LogLevel. Valid values: Trace, Debug, Information.

.PARAMETER SetEnumerationLogLevel
    Sets the BHE settings.json EnumerationLogLevel. Valid values: Trace, Debug, Information.

.PARAMETER RestartDelegatorAfterChange
    If specified along with -SetLogLevel and/or -SetEnumerationLogLevel, the script
    restarts the SHDelegator service automatically after applying the change(s).

.PARAMETER LogArchiveNumber
    Copy only N most recent files from the log_archive directory. If not specified,
    the entire log_archive directory is copied (if present).

.EXAMPLE
    .\GetBHESupportLogs.ps1

.EXAMPLE
    .\GetBHESupportLogs.ps1 -OutputRoot C:\Temp

.EXAMPLE
    # Set log levels via parameters and restart service automatically
    .\GetBHESupportLogs.ps1 -SetLogLevel Debug -SetEnumerationLogLevel Trace -RestartDelegatorAfterChange

.EXAMPLE
    # Collect with exclusions and limit log_archive files
    .\GetBHESupportLogs.ps1 -ExcludeEventLogs -ExcludeSettings -LogArchiveNumber 10

.NOTES
    - Windows PowerShell 5.1+ is supported (PowerShell 7+ also works)
    - Run as Administrator for best results (event log export and profile access)
    - If the SHDelegator service is not present, BHE files may not be found
    - No interactive prompts are shown for log levels or service restart
#>
[CmdletBinding()]
param(
    [string]$OutputRoot = "$env:USERPROFILE\Desktop",
    [string]$ServiceName = "SHDelegator",
    [switch]$AutoStart,
    [switch]$ExcludeEventLogs,
    [switch]$ExcludeSettings,
    [ValidateSet('Trace','Debug','Information')]
    [string]$SetLogLevel,
    [ValidateSet('Trace','Debug','Information')]
    [string]$SetEnumerationLogLevel,
    [switch]$RestartDelegatorAfterChange,
    [int]$LogArchiveNumber
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Info {
    param([string]$Message)
    Write-Host "[INFO] $Message"
}

function Write-Warn {
    param([string]$Message)
    Write-Warning $Message
}

# Simple ASCII banner and UI helpers
function Show-Banner {
    $label = 'BHE Logs Collector v.1.1'
    $width = [Math]::Max($label.Length + 8, 40)
    $border = ('=' * $width)
    $padLeft = [int][Math]::Floor(($width - $label.Length) / 2)
    $padRight = $width - $padLeft - $label.Length
    Clear-Host
    Write-Host $border -ForegroundColor Cyan
    Write-Host ((' ' * $padLeft) + $label + (' ' * $padRight)) -ForegroundColor Cyan
    Write-Host $border -ForegroundColor Cyan
}

function Wait-ForEnter {
    try {
        Write-Host "Press Enter to collect logs, (H)elp for parameters, or Q to quit"
        while ($true) {
            $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            if ($key.VirtualKeyCode -eq 13) { return $true }
            if ($key.Character -in @('q','Q')) { return $false }
            if ($key.Character -in @('h','H')) { 
                Show-CommandLineHelp
            }
        }
    } catch {
        $input = Read-Host
        if ($input -match '^[Qq]$') { return $false }
        if ($input -match '^[Hh]$') { 
            Show-CommandLineHelp
            return Wait-ForEnter
        }
        return $true
    }
}

function Show-CommandLineHelp {
    Write-Host ""; Write-Host "Command Line Parameters:" -ForegroundColor Cyan
    Write-Host "  -OutputRoot [path]           Output folder (default: Desktop)" -ForegroundColor White
    Write-Host "  -ServiceName [name]          Service name (default: SHDelegator)" -ForegroundColor White
    Write-Host "  -AutoStart                   Skip prompts, auto-collect" -ForegroundColor White
    Write-Host "  -ExcludeEventLogs            Skip Windows event logs" -ForegroundColor White
    Write-Host "  -ExcludeSettings             Skip settings.json" -ForegroundColor White
    Write-Host "  -SetLogLevel [level]         Set LogLevel (Trace|Debug|Information)" -ForegroundColor White
    Write-Host "  -SetEnumerationLogLevel [l]  Set EnumerationLogLevel (Trace|Debug|Information)" -ForegroundColor White
    Write-Host "  -RestartDelegatorAfterChange Auto-restart service after changes (no prompt)" -ForegroundColor White
    Write-Host "  -LogArchiveNumber [int]      Copy only N most recent files from log_archive" -ForegroundColor White
    Write-Host ""
    Write-Host "Examples:" -ForegroundColor Cyan
    Write-Host "  .\GetBHESupportLogsTool.ps1 -OutputRoot C:\Temp -AutoStart" -ForegroundColor DarkCyan
    Write-Host "  .\GetBHESupportLogsTool.ps1 -SetLogLevel Debug -SetEnumerationLogLevel Trace -RestartDelegatorAfterChange" -ForegroundColor DarkCyan
    Write-Host "  .\GetBHESupportLogsTool.ps1 -ExcludeEventLogs -ExcludeSettings" -ForegroundColor DarkCyan
    Write-Host "  .\GetBHESupportLogsTool.ps1 -LogArchiveNumber 10" -ForegroundColor DarkCyan
}

function Add-Status {
    param(
        [System.Collections.IList]$List,
        [string]$Name,
        [string]$Path,
        [string]$Status,
        [string]$Note = ''
    )
    $List.Add([pscustomobject]@{ Name = $Name; Path = $Path; Status = $Status; Note = $Note }) | Out-Null
}

function Print-ItemResult {
    param(
        [string]$Name,
        [string]$Status,
        [string]$Note = ''
    )
    $color = 'White'
    switch ($Status) {
        'Collected' { $color = 'Green' }
        'Created' { $color = 'Green' }
        'Updated' { $color = 'Green' }
        'Skipped' { $color = 'DarkYellow' }
        'NotFound' { $color = 'Yellow' }
        'Failed' { $color = 'Red' }
        default { $color = 'Gray' }
    }
    $noteText = if ([string]::IsNullOrWhiteSpace($Note)) { '' } else { " - $Note" }
    Write-Host ("  - $Name ... $Status$noteText") -ForegroundColor $color
}

<#
.SYNOPSIS
    Export Application and System Windows Event Logs and record status.
.DESCRIPTION
    Calls Export-EventLogs and then records which outputs were produced (EVTX or XML fallback).
#>
function Collect-EventLogsWithStatus {
    param(
        [string]$DestinationFolder,
        [System.Collections.IList]$StatusList
    )

    if ($script:ExcludeEventLogs) {
        Add-Status -List $StatusList -Name 'Application Event Log' -Path (Join-Path $DestinationFolder 'Application.evtx') -Status 'Skipped' -Note 'Excluded by parameter'
        Add-Status -List $StatusList -Name 'System Event Log' -Path (Join-Path $DestinationFolder 'System.evtx') -Status 'Skipped' -Note 'Excluded by parameter'
        Print-ItemResult -Name 'Application Event Log' -Status 'Skipped' -Note 'Excluded by parameter'
        Print-ItemResult -Name 'System Event Log' -Status 'Skipped' -Note 'Excluded by parameter'
        return
    }

    Write-Host "Collecting Windows Event Logs..." -ForegroundColor Cyan

    # Perform export using existing function
    Export-EventLogs -DestinationFolder $DestinationFolder

    # Application
    $appEvtx = Join-Path $DestinationFolder 'Application.evtx'
    $appXml  = Join-Path $DestinationFolder 'Application.xml'
    if (Test-Path -LiteralPath $appEvtx) {
        Add-Status -List $StatusList -Name 'Application Event Log' -Path $appEvtx -Status 'Collected' -Note 'EVTX'
        Print-ItemResult -Name 'Application Event Log' -Status 'Collected' -Note 'EVTX'
    } elseif (Test-Path -LiteralPath $appXml) {
        Add-Status -List $StatusList -Name 'Application Event Log' -Path $appXml -Status 'Collected' -Note 'XML fallback'
        Print-ItemResult -Name 'Application Event Log' -Status 'Collected' -Note 'XML fallback'
    } else {
        Add-Status -List $StatusList -Name 'Application Event Log' -Path $appEvtx -Status 'Failed' -Note 'Not exported'
        Print-ItemResult -Name 'Application Event Log' -Status 'Failed' -Note 'Not exported'
    }

    # System
    $sysEvtx = Join-Path $DestinationFolder 'System.evtx'
    $sysXml  = Join-Path $DestinationFolder 'System.xml'
    if (Test-Path -LiteralPath $sysEvtx) {
        Add-Status -List $StatusList -Name 'System Event Log' -Path $sysEvtx -Status 'Collected' -Note 'EVTX'
        Print-ItemResult -Name 'System Event Log' -Status 'Collected' -Note 'EVTX'
    } elseif (Test-Path -LiteralPath $sysXml) {
        Add-Status -List $StatusList -Name 'System Event Log' -Path $sysXml -Status 'Collected' -Note 'XML fallback'
        Print-ItemResult -Name 'System Event Log' -Status 'Collected' -Note 'XML fallback'
    } else {
        Add-Status -List $StatusList -Name 'System Event Log' -Path $sysEvtx -Status 'Failed' -Note 'Not exported'
        Print-ItemResult -Name 'System Event Log' -Status 'Failed' -Note 'Not exported'
    }
}

<#
.SYNOPSIS
    Copy BloodHound Enterprise artifacts from the service account profile and record status.
.DESCRIPTION
    Copies log_archive, service.log, and settings.json from the service account's
    Roaming AppData profile (falls back to current user if service profile not found).
#>
function Collect-BHEFilesWithStatus {
    param(
        [string]$ServiceProfilePath,
        [string]$DestinationFolder,
        [System.Collections.IList]$StatusList
    )

    Write-Host "Collecting BloodHoundEnterprise files..." -ForegroundColor Cyan

    $roaming = if ($ServiceProfilePath) { Join-Path $ServiceProfilePath 'AppData\Roaming' } else { $null }
    if (-not $roaming -or -not (Test-Path -LiteralPath $roaming)) {
        $roaming = [Environment]::GetFolderPath('ApplicationData')
    }
    $bheSource = Join-Path $roaming 'BloodHoundEnterprise'
    $bheOut = Join-Path $DestinationFolder 'BloodHoundEnterprise'
    New-Item -ItemType Directory -Path $bheOut -Force | Out-Null

    $items = @(
        @{ Name = 'BHE service.log'; Src = Join-Path $bheSource 'service.log'; Dest = Join-Path $bheOut 'service.log'; Dir = $false }
    )
    if (-not $script:ExcludeSettings) {
        $items += @{ Name = 'BHE settings.json'; Src = Join-Path $bheSource 'settings.json'; Dest = Join-Path $bheOut 'settings.json'; Dir = $false }
    } else {
        Add-Status -List $StatusList -Name 'BHE settings.json' -Path (Join-Path $bheSource 'settings.json') -Status 'Skipped' -Note 'Excluded by parameter'
        Print-ItemResult -Name 'BHE settings.json' -Status 'Skipped' -Note 'Excluded by parameter'
    }

    # Handle log_archive separately to support -LogArchiveNumber
    $logArchiveSrc = Join-Path $bheSource 'log_archive'
    $logArchiveDest = Join-Path $bheOut 'log_archive'
    if (Test-Path -LiteralPath $logArchiveSrc) {
        try {
            if (Test-Path -LiteralPath $logArchiveDest) {
                Remove-Item -LiteralPath $logArchiveDest -Recurse -Force -ErrorAction SilentlyContinue
            }
            New-Item -ItemType Directory -Path $logArchiveDest -Force | Out-Null

            if ($script:LogArchiveNumber -and $script:LogArchiveNumber -gt 0) {
                # Copy only the most recent N files
                $allFiles = Get-ChildItem -LiteralPath $logArchiveSrc -File | Sort-Object CreationTime -Descending
                $filesToCopy = $allFiles | Select-Object -First $script:LogArchiveNumber
                foreach ($file in $filesToCopy) {
                    Copy-Item -LiteralPath $file.FullName -Destination (Join-Path $logArchiveDest $file.Name) -Force -ErrorAction Stop
                }
                Add-Status -List $StatusList -Name 'BHE log_archive' -Path $logArchiveDest -Status 'Collected' -Note "Most recent $($filesToCopy.Count) of $($allFiles.Count) files"
                Print-ItemResult -Name 'BHE log_archive' -Status 'Collected' -Note "Most recent $($filesToCopy.Count) of $($allFiles.Count) files"
            } else {
                # Copy entire folder
                Copy-Item -LiteralPath $logArchiveSrc -Destination $logArchiveDest -Recurse -Force -ErrorAction Stop
                Add-Status -List $StatusList -Name 'BHE log_archive' -Path $logArchiveDest -Status 'Collected'
                Print-ItemResult -Name 'BHE log_archive' -Status 'Collected'
            }
        } catch {
            Add-Status -List $StatusList -Name 'BHE log_archive' -Path $logArchiveDest -Status 'Failed' -Note $($_.Exception.Message)
            Print-ItemResult -Name 'BHE log_archive' -Status 'Failed' -Note $($_.Exception.Message)
        }
    } else {
        Add-Status -List $StatusList -Name 'BHE log_archive' -Path $logArchiveSrc -Status 'NotFound'
        Print-ItemResult -Name 'BHE log_archive' -Status 'NotFound'
    }

    foreach ($it in $items) {
        $src = $it.Src
        $dest = $it.Dest
        $name = $it.Name

        if (Test-Path -LiteralPath $src) {
            try {
                if (Test-Path -LiteralPath $dest) {
                    Remove-Item -LiteralPath $dest -Recurse -Force -ErrorAction SilentlyContinue
                }
                if ($it.Dir) {
                    Copy-Item -LiteralPath $src -Destination $dest -Recurse -Force -ErrorAction Stop
                } else {
                    Copy-Item -LiteralPath $src -Destination $dest -Force -ErrorAction Stop
                }
                Add-Status -List $StatusList -Name $name -Path $dest -Status 'Collected'
                Print-ItemResult -Name $name -Status 'Collected'
            } catch {
                Add-Status -List $StatusList -Name $name -Path $dest -Status 'Failed' -Note $($_.Exception.Message)
                Print-ItemResult -Name $name -Status 'Failed' -Note $($_.Exception.Message)
            }
        } else {
            Add-Status -List $StatusList -Name $name -Path $src -Status 'NotFound'
            Print-ItemResult -Name $name -Status 'NotFound'
        }
    }
}

<#
.SYNOPSIS
    Print a one-screen summary of collected, missing, and failed items with paths.
#>
function Print-Summary {
    param(
        [System.Collections.IList]$StatusList,
        [string]$WorkDir,
        [string]$ZipPath
    )

    Write-Host ""; Write-Host "Summary" -ForegroundColor Cyan
    $ok = $StatusList | Where-Object { $_.Status -in @('Collected','Created','Updated') }
    $missing = $StatusList | Where-Object { $_.Status -eq 'NotFound' }
    $failed = $StatusList | Where-Object { $_.Status -eq 'Failed' }

    if ($ok) {
        Write-Host "Collected:" -ForegroundColor Green
        $ok | ForEach-Object { Write-Host ("  - {0} -> {1}" -f $_.Name, $_.Path) }
    } else {
        Write-Host "No items collected." -ForegroundColor Yellow
    }
    if ($missing -or $failed) {
        Write-Host "Issues:" -ForegroundColor Yellow
        $missing | ForEach-Object { Write-Host ("  - {0}: Not found ({1})" -f $_.Name, $_.Path) -ForegroundColor Yellow }
        $failed | ForEach-Object { Write-Host ("  - {0}: Failed - {1}" -f $_.Name, $_.Note) -ForegroundColor Red }
    }

    Write-Host ""; Write-Host ("Output folder: {0}" -f $WorkDir)
    Write-Host ("Zip archive:  {0}" -f $ZipPath)
    try {
        $workUri = (New-Object System.Uri($WorkDir)).AbsoluteUri
        $zipUri = (New-Object System.Uri($ZipPath)).AbsoluteUri
        Write-Host ("Open folder: {0}" -f $workUri) -ForegroundColor DarkCyan
        Write-Host ("Open zip:    {0}" -f $zipUri) -ForegroundColor DarkCyan
    } catch { }
}

<#
.SYNOPSIS
    Offer to open the output folder or select the created zip in Explorer.
#>
function Prompt-Open {
    param(
        [string]$WorkDir,
        [string]$ZipPath
    )
    Write-Host ""; Write-Host "Press O to open output folder, Z to open at zip, or any other key to exit." -ForegroundColor Cyan
    $choice = Read-Host "Choice"
    switch ($choice.ToUpperInvariant()) {
        'O' { Start-Process explorer.exe $WorkDir }
        'Z' { Start-Process explorer.exe "/select,`"$ZipPath`"" }
        default { }
    }
}

<#
.SYNOPSIS
    Resolve the SHDelegator service object using exact matches.
.DESCRIPTION
    Attempts to locate the service by Name, then DisplayName, then Description. Returns $null if not found.
#>
function Get-ServiceObject {
    param([string]$Name)

    # Try exact service name first (e.g., SHDelegator)
    try {
        $svc = Get-CimInstance -ClassName Win32_Service -Filter "Name='$Name'" -ErrorAction Stop
        if ($null -ne $svc) { return ($svc | Select-Object -First 1) }
    } catch { }

    # Try exact display name (SharpHoundDelegator)
    try {
        $svc = Get-CimInstance -ClassName Win32_Service -Filter "DisplayName='SharpHoundDelegator'" -ErrorAction Stop
        if ($null -ne $svc) { return ($svc | Select-Object -First 1) }
    } catch { }

    # Try exact description (SharpHound Delegation Service)
    try {
        $svc = Get-CimInstance -ClassName Win32_Service -Filter "Description='SharpHound Delegation Service'" -ErrorAction Stop
        if ($null -ne $svc) { return ($svc | Select-Object -First 1) }
    } catch { }

    Write-Warn "Service not found by name '$Name', display name 'SharpHoundDelegator', or description 'SharpHound Delegation Service'."
    Write-Warn "Proceeding without service context; current user's %AppData% will be used for BHE files if present."
    return $null
}

<#
.SYNOPSIS
    Compute the profile path for a service account given its StartName.
.DESCRIPTION
    Handles built-in service accounts (LocalSystem, LocalService, NetworkService) and
    domain/local users by translating the NTAccount to a SID and reading ProfileImagePath.
#>
function Get-ProfilePathFromServiceStartName {
    param([string]$StartName)

    if ([string]::IsNullOrWhiteSpace($StartName)) { return $null }

    $normalized = $StartName.ToLowerInvariant()
    switch ($normalized) {
        'localsystem' { return Join-Path $env:WINDIR 'System32\config\systemprofile' }
        'nt authority\system' { return Join-Path $env:WINDIR 'System32\config\systemprofile' }
        'localservice' { return Join-Path $env:WINDIR 'ServiceProfiles\LocalService' }
        'nt authority\localservice' { return Join-Path $env:WINDIR 'ServiceProfiles\LocalService' }
        'networkservice' { return Join-Path $env:WINDIR 'ServiceProfiles\NetworkService' }
        'nt authority\networkservice' { return Join-Path $env:WINDIR 'ServiceProfiles\NetworkService' }
        default {
            try {
                $ntAccount = New-Object System.Security.Principal.NTAccount($StartName)
                $sid = $ntAccount.Translate([System.Security.Principal.SecurityIdentifier]).Value
                $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid"
                $profileImagePath = (Get-ItemProperty -Path $regPath -Name 'ProfileImagePath' -ErrorAction Stop).ProfileImagePath
                $expandedPath = [System.Environment]::ExpandEnvironmentVariables($profileImagePath)
                return $expandedPath
            } catch {
                Write-Warn "Failed to resolve profile path for service account '$StartName': $($_.Exception.Message)"
                return $null
            }
        }
    }
}

<#
.SYNOPSIS
    Export Application and System logs to EVTX (or XML fallback) using wevtutil/Get-WinEvent.
#>
function Export-EventLogs {
    param([string]$DestinationFolder)

    $appOut = Join-Path $DestinationFolder 'Application.evtx'
    $sysOut = Join-Path $DestinationFolder 'System.evtx'

    Write-Info 'Exporting Application and System event logs...'
    try {
        & wevtutil epl Application $appOut /ow:true | Out-Null
    } catch {
        Write-Warn "Failed to export Application log via wevtutil: $($_.Exception.Message)"
        try {
            Get-WinEvent -LogName Application -ErrorAction Stop | Export-Clixml -Path (Join-Path $DestinationFolder 'Application.xml')
            Write-Warn 'Exported Application log as XML fallback.'
        } catch {
            Write-Warn "Failed to export Application log fallback: $($_.Exception.Message)"
        }
    }

    try {
        & wevtutil epl System $sysOut /ow:true | Out-Null
    } catch {
        Write-Warn "Failed to export System log via wevtutil: $($_.Exception.Message)"
        try {
            Get-WinEvent -LogName System -ErrorAction Stop | Export-Clixml -Path (Join-Path $DestinationFolder 'System.xml')
            Write-Warn 'Exported System log as XML fallback.'
        } catch {
            Write-Warn "Failed to export System log fallback: $($_.Exception.Message)"
        }
    }
}

<#
.SYNOPSIS
    Copy core BloodHound Enterprise files given a known profile path.
.DESCRIPTION
    Used by the status-collecting wrapper; left available for reuse and scripting.
#>
function Copy-BHEFiles {
    param(
        [string]$ProfilePath,
        [string]$DestinationFolder
    )

    $roaming = if ($ProfilePath) { Join-Path $ProfilePath 'AppData\Roaming' } else { $null }
    if (-not $roaming -or -not (Test-Path -LiteralPath $roaming)) {
        Write-Warn 'Roaming AppData path not found; attempting current user AppData.'
        $roaming = [Environment]::GetFolderPath('ApplicationData')
    }

    $bheSource = Join-Path $roaming 'BloodHoundEnterprise'
    if (-not (Test-Path -LiteralPath $bheSource)) {
        Write-Warn "BloodHoundEnterprise folder not found at '$bheSource'. Skipping file collection from that location."
        return
    }

    $bheOut = Join-Path $DestinationFolder 'BloodHoundEnterprise'
    New-Item -ItemType Directory -Path $bheOut -Force | Out-Null

    $itemsToCopy = @(
        @{ Path = Join-Path $bheSource 'log_archive'; Dest = Join-Path $bheOut 'log_archive'; Recurse = $true },
        @{ Path = Join-Path $bheSource 'service.log'; Dest = Join-Path $bheOut 'service.log'; Recurse = $false },
        @{ Path = Join-Path $bheSource 'settings.json'; Dest = Join-Path $bheOut 'settings.json'; Recurse = $false }
    )

    foreach ($item in $itemsToCopy) {
        $src = $item.Path
        $dest = $item.Dest
        $recurse = [bool]$item.Recurse
        if (Test-Path -LiteralPath $src) {
            try {
                if (Test-Path -LiteralPath $dest) {
                    Remove-Item -LiteralPath $dest -Recurse -Force -ErrorAction SilentlyContinue
                }
                if ($recurse) {
                    Copy-Item -LiteralPath $src -Destination $dest -Recurse -Force -ErrorAction Stop
                } else {
                    Copy-Item -LiteralPath $src -Destination $dest -Force -ErrorAction Stop
                }
                Write-Info "Copied '$src' -> '$dest'"
            } catch {
                Write-Warn "Failed to copy '$src' -> '$dest': $($_.Exception.Message)"
            }
        } else {
            Write-Warn "Source not found, skipping: '$src'"
        }
    }
}

# NOTE: Interactive prompts for LogLevel and EnumerationLogLevel have been removed.
# Changes are applied only when -SetLogLevel and/or -SetEnumerationLogLevel are provided.
function Set-BHELogLevel {
    param(
        [string]$ServiceProfilePath,
        [string]$DesiredLevel,
        [System.Collections.IList]$StatusList
    )

    if ([string]::IsNullOrWhiteSpace($DesiredLevel)) { return }

    if ($script:ExcludeSettings) {
        Add-Status -List $StatusList -Name 'BHE settings.json LogLevel' -Path '' -Status 'Skipped' -Note 'Excluded by parameter'
        Print-ItemResult -Name 'BHE settings.json LogLevel' -Status 'Skipped' -Note 'Excluded by parameter'
        return
    }

    $roaming = if ($ServiceProfilePath) { Join-Path $ServiceProfilePath 'AppData\Roaming' } else { [Environment]::GetFolderPath('ApplicationData') }
    $settingsPath = Join-Path (Join-Path $roaming 'BloodHoundEnterprise') 'settings.json'

    if (-not (Test-Path -LiteralPath $settingsPath)) {
        Add-Status -List $StatusList -Name 'BHE settings.json LogLevel' -Path $settingsPath -Status 'NotFound'
        Print-ItemResult -Name 'BHE settings.json LogLevel' -Status 'NotFound'
        return
    }

    $backupPath = $settingsPath + ('.bak_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
    try {
        Copy-Item -LiteralPath $settingsPath -Destination $backupPath -Force -ErrorAction Stop
    } catch { }

    try {
        $json = Get-Content -LiteralPath $settingsPath -Raw | ConvertFrom-Json -ErrorAction Stop
        if ($null -eq $json) { throw 'settings.json parsed as null' }
        $json.LogLevel = $DesiredLevel
        $json | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $settingsPath -Encoding UTF8 -Force
        Add-Status -List $StatusList -Name 'BHE settings.json LogLevel' -Path $settingsPath -Status 'Updated' -Note $DesiredLevel
        Print-ItemResult -Name 'BHE settings.json LogLevel' -Status 'Updated' -Note $DesiredLevel
    } catch {
        Add-Status -List $StatusList -Name 'BHE settings.json LogLevel' -Path $settingsPath -Status 'Failed' -Note $($_.Exception.Message)
        Print-ItemResult -Name 'BHE settings.json LogLevel' -Status 'Failed' -Note $($_.Exception.Message)
    }
}

function Set-BHEEnumerationLogLevel {
    param(
        [string]$ServiceProfilePath,
        [string]$DesiredLevel,
        [System.Collections.IList]$StatusList
    )

    if ([string]::IsNullOrWhiteSpace($DesiredLevel)) { return }

    if ($script:ExcludeSettings) {
        Add-Status -List $StatusList -Name 'BHE settings.json EnumerationLogLevel' -Path '' -Status 'Skipped' -Note 'Excluded by parameter'
        Print-ItemResult -Name 'BHE settings.json EnumerationLogLevel' -Status 'Skipped' -Note 'Excluded by parameter'
        return
    }

    $roaming = if ($ServiceProfilePath) { Join-Path $ServiceProfilePath 'AppData\Roaming' } else { [Environment]::GetFolderPath('ApplicationData') }
    $settingsPath = Join-Path (Join-Path $roaming 'BloodHoundEnterprise') 'settings.json'

    if (-not (Test-Path -LiteralPath $settingsPath)) {
        Add-Status -List $StatusList -Name 'BHE settings.json EnumerationLogLevel' -Path $settingsPath -Status 'NotFound'
        Print-ItemResult -Name 'BHE settings.json EnumerationLogLevel' -Status 'NotFound'
        return
    }

    $backupPath = $settingsPath + ('.bak_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
    try {
        Copy-Item -LiteralPath $settingsPath -Destination $backupPath -Force -ErrorAction Stop
    } catch { }

    try {
        $json = Get-Content -LiteralPath $settingsPath -Raw | ConvertFrom-Json -ErrorAction Stop
        if ($null -eq $json) { throw 'settings.json parsed as null' }
        $json.EnumerationLogLevel = $DesiredLevel
        $json | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $settingsPath -Encoding UTF8 -Force
        Add-Status -List $StatusList -Name 'BHE settings.json EnumerationLogLevel' -Path $settingsPath -Status 'Updated' -Note $DesiredLevel
        Print-ItemResult -Name 'BHE settings.json EnumerationLogLevel' -Status 'Updated' -Note $DesiredLevel
    } catch {
        Add-Status -List $StatusList -Name 'BHE settings.json EnumerationLogLevel' -Path $settingsPath -Status 'Failed' -Note $($_.Exception.Message)
        Print-ItemResult -Name 'BHE settings.json EnumerationLogLevel' -Status 'Failed' -Note $($_.Exception.Message)
    }
}

function Try-RestartDelegatorService {
    param([string]$Name)
    try {
        $svc = Get-Service -Name $Name -ErrorAction Stop
    } catch {
        try { $svc = Get-Service -DisplayName 'SharpHoundDelegator' -ErrorAction Stop } catch { $svc = $null }
    }
    if (-not $svc) { Write-Warn "Delegator service not found for restart."; return }

    try {
        Write-Info "Restarting service '$($svc.Name)'..."
        Restart-Service -InputObject $svc -ErrorAction Stop
        Write-Info "Service restarted."
    } catch {
        Write-Warn "Failed to restart service '$($svc.Name)': $($_.Exception.Message)"
    }
}

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$workDir = Join-Path $OutputRoot "BHE_SupportLogs_$timestamp"
$zipPath = Join-Path $OutputRoot "BHE_SupportLogs_$timestamp.zip"

$transcriptPath = Join-Path $workDir 'collectorlogs.log'

Show-Banner

# Warn about potentially sensitive content
Write-Warning 'Note: This collection will include the below data!'
if (-not $ExcludeEventLogs) { Write-Warning 'Windows Application and System event logs will be collected; use -ExcludeEventLogs to skip.' }
if (-not $ExcludeSettings) { Write-Warning 'settings.json will be collected; use -ExcludeSettings to skip.' }

# Validate OutputRoot exists and is writable
try {
    if (-not (Test-Path -LiteralPath $OutputRoot)) {
        throw "OutputRoot not found: $OutputRoot"
    }
    $testFile = Join-Path $OutputRoot ("write_test_" + [Guid]::NewGuid().ToString() + ".tmp")
    Set-Content -LiteralPath $testFile -Value 'test' -Encoding ascii -Force
    Remove-Item -LiteralPath $testFile -Force -ErrorAction SilentlyContinue
} catch {
    Write-Error "OutputRoot validation failed: $($_.Exception.Message)"
    return
}

$interactive = $Host.UI -and $Host.UI.RawUI -and -not $AutoStart.IsPresent -and $PSBoundParameters.ContainsKey('AutoStart') -eq $false
if ($interactive) {
    if (-not (Wait-ForEnter)) {
        Write-Host "Exiting by user request." -ForegroundColor Yellow
        return
    }
}

try {
    try {
        New-Item -ItemType Directory -Path $workDir -Force | Out-Null
    } catch {
        Write-Error "Failed to create output directory '$workDir': $($_.Exception.Message)"
        exit 1
    }

    try { Start-Transcript -Path $transcriptPath -ErrorAction SilentlyContinue | Out-Null } catch { }

    Write-Info "Output folder: $workDir"

    # Status accumulator used across collection stages
    $records = New-Object System.Collections.ArrayList

    # Resolve service and profile path (for settings.json and file collection)
    $svc = Get-ServiceObject -Name $ServiceName
    $profilePath = $null
    if ($svc) {
        Write-Info "Using service '$($svc.Name)' (DisplayName: '$($svc.DisplayName)') running as '$($svc.StartName)'"
        $profilePath = Get-ProfilePathFromServiceStartName -StartName $svc.StartName
        if ($profilePath) {
            Write-Info "Resolved service profile path: $profilePath"
        } else {
            Write-Warn 'Could not resolve service profile path; will attempt current user profile.'
        }
    }

    # Parameter-only: set BHE LogLevel and EnumerationLogLevel
    $desiredLevel = $SetLogLevel
    $desiredEnumLevel = $SetEnumerationLogLevel

    $didChange = $false
    if ($desiredLevel) {
        Set-BHELogLevel -ServiceProfilePath $profilePath -DesiredLevel $desiredLevel -StatusList $records
        $didChange = $true
    }
    if ($desiredEnumLevel) {
        Set-BHEEnumerationLogLevel -ServiceProfilePath $profilePath -DesiredLevel $desiredEnumLevel -StatusList $records
        $didChange = $true
    }

    # Parameter-only restart (no prompt)
    if ($didChange -and $RestartDelegatorAfterChange) {
        Try-RestartDelegatorService -Name $ServiceName
    }

    # Event logs
    Collect-EventLogsWithStatus -DestinationFolder $workDir -StatusList $records

    # BHE files
    Collect-BHEFilesWithStatus -ServiceProfilePath $profilePath -DestinationFolder $workDir -StatusList $records

    Write-Host ''
    try { Stop-Transcript | Out-Null } catch { }
    Write-Info 'Creating zip archive...'
    if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue }
    Compress-Archive -Path (Join-Path $workDir '*') -DestinationPath $zipPath -Force
    Add-Status -List $records -Name 'Zip Archive' -Path $zipPath -Status 'Created'
    Print-ItemResult -Name 'Zip Archive' -Status 'Created'

    Write-Host "\nCollection complete." -ForegroundColor Green
    Write-Host "Folder: $workDir"
    Write-Host "Zip:    $zipPath"

    Print-Summary -StatusList $records -WorkDir $workDir -ZipPath $zipPath
    if ($interactive) { Prompt-Open -WorkDir $workDir -ZipPath $zipPath }
} catch {
    Write-Error "Unexpected error: $($_.Exception.Message)"
} finally {
    try { Stop-Transcript | Out-Null } catch { }
}
