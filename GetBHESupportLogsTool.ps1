<#
.SYNOPSIS
    Collects logs and configuration relevant to BloodHound Enterprise troubleshooting.

.DESCRIPTION
    This tool exports Windows Application and System event logs, collects
    BloodHound Enterprise log files and settings from the profile of the account
    running the SHDelegator service, and additionally can also collect AzureHound logs, then compresses the results into a zip archive.

    It displays a simple text UI with progress for each item, a summary of collected
    and missing items, and options to open the output folder or the zip file.

    The tool can collect SharpHound logs, AzureHound logs, or both using the -All
    parameter. When using -All, all logs are collected regardless of interactive
    target selection.

    Additionally, the tool can collect a Performance Monitor trace and zip the output.

    NOTE: Log level changes (e.g., LogLevel and EnumerationLogLevel) and service restarts
    are controlled only via command-line parameters; there are no interactive prompts
    for these settings.


.EXAMPLE
    .\GetBHESupportLogsTool.ps1

.EXAMPLE
    .\GetBHESupportLogsTool.ps1 -OutputRoot "C:\Temp"

.EXAMPLE
    # Collect all logs (SharpHound, AzureHound, and event logs)
    .\GetBHESupportLogsTool.ps1 -All

.EXAMPLE
    # Set log levels via parameters and restart service explicitly
    .\GetBHESupportLogsTool.ps1 -SetLogLevel Debug -SetEnumerationLogLevel Trace -RestartDelegator

.EXAMPLE
    # Collect with exclusions and limit log_archive files
    .\GetBHESupportLogsTool.ps1 -ExcludeEventLogs -ExcludeSettings -LogArchiveNumber 5

.EXAMPLE
    # Show command line help
    .\GetBHESupportLogsTool.ps1 -Help

.NOTES
    - Windows PowerShell 5.1+ is supported (PowerShell 7+ also works)
    - Run as Administrator for best results (event log export and profile access)
    - No interactive prompts are shown for log levels or service restart
    - See README.md for more information
#>
[CmdletBinding()]
param(
    [string]$OutputRoot = "$env:USERPROFILE\Desktop",
    [switch]$ExcludeEventLogs,
    [switch]$ExcludeSettings,
    [ValidateSet('Trace','Debug','Information')]
    [string]$SetLogLevel,
    [ValidateSet('Trace','Debug','Information')]
    [string]$SetEnumerationLogLevel,
    [switch]$RestartDelegator,
    [int]$LogArchiveNumber,
    [ValidateSet(0,1,2)]
    [int]$SetAzureVerbosity,
    [switch]$RestartAzureHound,
    [switch]$All,
    [switch]$AllPlusPerf,
    [switch]$GetBHEPerfmon,
    [switch]$DeleteBHEPerfmon,
    [string]$GetCompStatus,
    [switch]$Help
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
    $label = 'BHE Logs Collector v.2.10'
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
        $loopCount = 0
        while ($true) {
            $loopCount++
            if ($loopCount -gt 10) {
                Write-Host "Too many loops detected, exiting..." -ForegroundColor Red
                return $false
            }
            Write-Host "Press Enter to collect logs, or Q to quit"
            $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            if ($key.VirtualKeyCode -eq 13) { return $true }
            if ($key.Character -in @('q','Q')) { return $false }
            if ($key.Character -in @('h','H')) { 
                Show-CommandLineHelp
                Write-Host ""
                # Continue the loop - will show prompt again
            }
        }
    } catch {
        $loopCount = 0
        while ($true) {
            $loopCount++
            if ($loopCount -gt 10) {
                Write-Host "Too many loops detected, exiting..." -ForegroundColor Red
                return $false
            }

            $input = Read-Host
            if ($input -match '^[Qq]$') { return $false }
            if ($input -match '^[Hh]$') { 
                Show-CommandLineHelp
                Write-Host ""
                # Continue the loop - will show prompt again
            } else {
                return $true
            }
        }
    }
}

function Show-CommandLineHelp {
    Write-Host ""; Write-Host "Command Line Parameters:" -ForegroundColor Cyan
    Write-Host "  -Help                            Show this help information" -ForegroundColor White
    Write-Host "  -OutputRoot [path]               Output folder (default: Desktop)" -ForegroundColor White
    Write-Host "  -All                             Collect all SharpHound, AzureHound, and event logs" -ForegroundColor White
    Write-Host "  -AllPlusPerf                     Do -All plus set up perf tracing" -ForegroundColor White
    Write-Host "  -GetBHEPerfmon                   Only set up/start BHE perfmon trace" -ForegroundColor White
    Write-Host "  -DeleteBHEPerfmon                Delete the BHE perfmon collector and logs" -ForegroundColor White
    Write-Host "  -GetCompStatus [path]            Analyze compstatus.csv file at specified path" -ForegroundColor White
    Write-Host "  -ExcludeEventLogs                Skip Windows event logs" -ForegroundColor White
    Write-Host "  -ExcludeSettings                 Skip settings.json" -ForegroundColor White
    Write-Host "  -SetLogLevel [level]             Set LogLevel (Trace|Debug|Information)" -ForegroundColor White
    Write-Host "  -SetEnumerationLogLevel [level]  Set EnumerationLogLevel (Trace|Debug|Information)" -ForegroundColor White
    Write-Host "  -RestartDelegator                Restart SHDelegator service without changing settings" -ForegroundColor White
    Write-Host "  -LogArchiveNumber [int]          Copy only N most recent files from log_archive" -ForegroundColor White
    Write-Host "  -SetAzureVerbosity [0|1|2]       Set AzureHound verbosity (0=Default,1=Debug,2=Trace)" -ForegroundColor White
    Write-Host "  -RestartAzureHound               Restart AzureHound service (with or without -SetAzureVerbosity)" -ForegroundColor White
    Write-Host ""
    Write-Host "Examples:" -ForegroundColor Cyan
    Write-Host "  .\GetBHESupportLogsTool.ps1 -Help" -ForegroundColor DarkCyan
    Write-Host "  .\GetBHESupportLogsTool.ps1 -OutputRoot 'C:\Temp'" -ForegroundColor DarkCyan
    Write-Host "  .\GetBHESupportLogsTool.ps1 -All" -ForegroundColor DarkCyan
    Write-Host "  .\GetBHESupportLogsTool.ps1 -AllPlusPerf" -ForegroundColor DarkCyan
    Write-Host "  .\GetBHESupportLogsTool.ps1 -GetBHEPerfmon" -ForegroundColor DarkCyan
    Write-Host "  .\GetBHESupportLogsTool.ps1 -DeleteBHEPerfmon" -ForegroundColor DarkCyan
    Write-Host "  .\GetBHESupportLogsTool.ps1 -SetLogLevel Debug -SetEnumerationLogLevel Trace -RestartDelegator" -ForegroundColor DarkCyan
    Write-Host "  .\GetBHESupportLogsTool.ps1 -ExcludeEventLogs -ExcludeSettings" -ForegroundColor DarkCyan
    Write-Host "  .\GetBHESupportLogsTool.ps1 -LogArchiveNumber 10" -ForegroundColor DarkCyan
    Write-Host "  .\GetBHESupportLogsTool.ps1 -SetAzureVerbosity 2 -RestartAzureHound" -ForegroundColor DarkCyan
    Write-Host "  .\GetBHESupportLogsTool.ps1 -RestartDelegator" -ForegroundColor DarkCyan
    Write-Host "  .\GetBHESupportLogsTool.ps1 -GetCompStatus 'C:\path\to\compstatus.csv'" -ForegroundColor DarkCyan
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
                $allFiles = @(Get-ChildItem -LiteralPath $logArchiveSrc -File | Sort-Object CreationTime -Descending)
                $filesToCopy = @($allFiles | Select-Object -First $script:LogArchiveNumber)
                foreach ($file in $filesToCopy) {
                    Copy-Item -LiteralPath $file.FullName -Destination (Join-Path $logArchiveDest $file.Name) -Force -ErrorAction Stop
                }
                Add-Status -List $StatusList -Name 'BHE log_archive' -Path $logArchiveDest -Status 'Collected' -Note "Most recent $($filesToCopy.Count) of $($allFiles.Count) files"
                Print-ItemResult -Name 'BHE log_archive' -Status 'Collected' -Note "Most recent $($filesToCopy.Count) of $($allFiles.Count) files"
            } else {
                # Copy entire folder
                Get-ChildItem -LiteralPath $logArchiveSrc | Copy-Item -Destination $logArchiveDest -Recurse -Force -ErrorAction Stop
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

    Write-Host ""; Write-Host ("Output folder: {0}" -f $WorkDir) -ForegroundColor DarkCyan
    Write-Host ("Zip archive:  {0}" -f $ZipPath) -ForegroundColor DarkCyan
    try {
        $workUri = (New-Object System.Uri($WorkDir)).AbsoluteUri
        $zipUri = (New-Object System.Uri($ZipPath)).AbsoluteUri
        #Write-Host ("Open folder: {0}" -f $workUri) -ForegroundColor DarkCyan
        #Write-Host ("Open zip:    {0}" -f $zipUri) -ForegroundColor DarkCyan
    } catch { }
}

function Prompt-SelectTarget {
    Write-Host ""; Write-Host "Select collection target: (S)harpHound or (A)zureHound" -ForegroundColor Cyan
    $choice = Read-Host "Choice [S/A]"
    switch ($choice.ToUpperInvariant()) {
        'A' { return 'AzureHound' }
        default { return 'SharpHound' }
    }
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

function Get-AzureHoundServiceObject {
    try {
        $svc = Get-CimInstance -ClassName Win32_Service -Filter "Name='AzureHound'" -ErrorAction Stop
        if ($null -ne $svc) { return ($svc | Select-Object -First 1) }
    } catch { }
    try {
        $svc = Get-CimInstance -ClassName Win32_Service -Filter "DisplayName='AzureHound'" -ErrorAction Stop
        if ($null -ne $svc) { return ($svc | Select-Object -First 1) }
    } catch { }
    try {
        $svc = Get-CimInstance -ClassName Win32_Service -Filter "Description='The official tool for collecting Azure data for BloodHound and BloodHound Enterprise.'" -ErrorAction Stop
        if ($null -ne $svc) { return ($svc | Select-Object -First 1) }
    } catch { }
    Write-Warn "AzureHound service not found by name/display/description."
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
        # Wait up to 30 seconds for the service to report Running
        $timeout = [TimeSpan]::FromSeconds(30)
        try {
            $svc.WaitForStatus('Running', $timeout)
        } catch { }
        $svc.Refresh()
        if ($svc.Status -eq 'Running') {
            Write-Info "Service '$($svc.Name)' is Running. Restart verified."
        } else {
            Write-Warn "Service '$($svc.Name)' did not reach Running state within $($timeout.TotalSeconds) seconds (current: $($svc.Status))."
        }
    } catch {
        Write-Warn "Failed to restart service '$($svc.Name)': $($_.Exception.Message)"
    }
}

function Collect-AzureHoundFilesWithStatus {
    param(
        [string]$DestinationFolder,
        [System.Collections.IList]$StatusList
    )

    Write-Host "Collecting AzureHound files..." -ForegroundColor Cyan
    $azOut = Join-Path $DestinationFolder 'AzureHound'
    New-Item -ItemType Directory -Path $azOut -Force | Out-Null

    $azLogSrc = 'C:\\Program Files\\AzureHound Enterprise\\azurehound.log'
    $azLogDest = Join-Path $azOut 'azurehound.log'
    if (Test-Path -LiteralPath $azLogSrc) {
        try {
            Copy-Item -LiteralPath $azLogSrc -Destination $azLogDest -Force -ErrorAction Stop
            Add-Status -List $StatusList -Name 'AzureHound azurehound.log' -Path $azLogDest -Status 'Collected'
            Print-ItemResult -Name 'AzureHound azurehound.log' -Status 'Collected'
        } catch {
            Add-Status -List $StatusList -Name 'AzureHound azurehound.log' -Path $azLogDest -Status 'Failed' -Note $($_.Exception.Message)
            Print-ItemResult -Name 'AzureHound azurehound.log' -Status 'Failed' -Note $($_.Exception.Message)
        }
    } else {
        Add-Status -List $StatusList -Name 'AzureHound azurehound.log' -Path $azLogSrc -Status 'NotFound'
        Print-ItemResult -Name 'AzureHound azurehound.log' -Status 'NotFound'
    }
}

function Set-AzureHoundVerbosity {
    param(
        [int]$Verbosity,
        [System.Collections.IList]$StatusList
    )

    if ($PSBoundParameters.ContainsKey('Verbosity') -eq $false) { return }
    $cfgPath = 'C:\\ProgramData\\azurehound\\config.json'
    if (-not (Test-Path -LiteralPath $cfgPath)) {
        Add-Status -List $StatusList -Name 'AzureHound config.json verbosity' -Path $cfgPath -Status 'NotFound'
        Print-ItemResult -Name 'AzureHound config.json verbosity' -Status 'NotFound'
        return
    }
    $backupPath = $cfgPath + ('.bak_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
    try { Copy-Item -LiteralPath $cfgPath -Destination $backupPath -Force -ErrorAction Stop } catch { }
    try {
        $json = Get-Content -LiteralPath $cfgPath -Raw | ConvertFrom-Json -ErrorAction Stop
        if ($null -eq $json) { throw 'config.json parsed as null' }
        $json.verbosity = [int]$Verbosity
        $json | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $cfgPath -Encoding UTF8 -Force
        Add-Status -List $StatusList -Name 'AzureHound config.json verbosity' -Path $cfgPath -Status 'Updated' -Note $Verbosity
        Print-ItemResult -Name 'AzureHound config.json verbosity' -Status 'Updated' -Note $Verbosity
    } catch {
        Add-Status -List $StatusList -Name 'AzureHound config.json verbosity' -Path $cfgPath -Status 'Failed' -Note $($_.Exception.Message)
        Print-ItemResult -Name 'AzureHound config.json verbosity' -Status 'Failed' -Note $($_.Exception.Message)
    }
}

function Try-RestartAzureHoundService {
    try {
        $svc = Get-Service -Name 'AzureHound' -ErrorAction Stop
    } catch {
        try { $svc = Get-Service -DisplayName 'AzureHound' -ErrorAction Stop } catch { $svc = $null }
    }
    if (-not $svc) { Write-Warn "AzureHound service not found for restart."; return }
    try {
        Write-Info "Restarting service '$($svc.Name)'..."
        Restart-Service -InputObject $svc -ErrorAction Stop
        # Wait up to 30 seconds for the service to report Running
        $timeout = [TimeSpan]::FromSeconds(30)
        try {
            $svc.WaitForStatus('Running', $timeout)
        } catch { }
        $svc.Refresh()
        if ($svc.Status -eq 'Running') {
            Write-Info "Service '$($svc.Name)' is Running. Restart verified."
        } else {
            Write-Warn "Service '$($svc.Name)' did not reach Running state within $($timeout.TotalSeconds) seconds (current: $($svc.Status))."
        }
    } catch {
        Write-Warn "Failed to restart service '$($svc.Name)': $($_.Exception.Message)"
    }
}

# ========================
# CompStatus Analysis Section
# ========================

function Invoke-CompStatusAnalysis {
    param(
        [string]$CsvPath
    )
    
    if (-not (Test-Path -LiteralPath $CsvPath)) {
        Write-Error "CompStatus CSV file not found: $CsvPath"
        return
    }
    
    Write-Host "Analyzing CompStatus data from: $CsvPath" -ForegroundColor Cyan
    Write-Host ""
    
    try {
        # Import data and get uniques without sorting them
        $stats_file = Import-Csv -Path $CsvPath | Group-Object ComputerName, Task, Status, IPAddress | ForEach-Object { $_.Group[0] }
        
        Write-Host "=== Status Pivot Table (Excluding GetMembersInAlias) ===" -ForegroundColor Yellow
        $stats_file | Where-Object {$_.Task -NotLike 'GetMembersInAlias -*'} | Group-Object Task, Status -NoElement | Format-Table -Autosize
        
        Write-Host ""
        Write-Host "=== Failures Only ===" -ForegroundColor Yellow
        $stats_file | Where-Object {$_.Status -ne "Success"} | Group-Object Task,Status -NoElement | Format-Table -Autosize
        
        Write-Host ""
        Write-Host "=== Systems Unreachable on 445/TCP ===" -ForegroundColor Yellow
        $unreachable = $stats_file | Where-Object {$_.Task -eq "ComputerAvailability" -and $_.Status -eq "PortNotOpen"}
        if ($unreachable) {
            $unreachable | Format-Table -Autosize
        } else {
            Write-Host "No systems unreachable on port 445/TCP" -ForegroundColor Green
        }
        
        Write-Host ""
        Write-Host "=== IPv4 /24 Subnets Unreachable on 445/TCP ===" -ForegroundColor Yellow
        $ipv4Unreachable = $stats_file | Where-Object {$_.Task -eq "ComputerAvailability" -and $_.Status -eq "PortNotOpen" -and $_.IPAddress -match '^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$'}
        if ($ipv4Unreachable) {
            $ipv4Unreachable | Group-Object {$_.IPAddress.Remove($_.IPAddress.LastIndexOf('.'))+'.0/24'} -NoElement | Sort-Object -Property Count | Format-Table -Autosize
        } else {
            Write-Host "No IPv4 systems unreachable on port 445/TCP" -ForegroundColor Green
        }
        
        Write-Host ""
        Write-Host "=== IPv4 /16 Subnets Unreachable on 445/TCP ===" -ForegroundColor Yellow
        if ($ipv4Unreachable) {
            $ipv4Unreachable | Group-Object {($_.IPAddress.split(".")[0..1] -join ".") + ".0.0/16"} -NoElement | Sort-Object -Property Count | Format-Table -Autosize
        } else {
            Write-Host "No IPv4 systems unreachable on port 445/TCP" -ForegroundColor Green
        }
        
        Write-Host ""
        Write-Host "=== Systems Missing Permissions ===" -ForegroundColor Yellow
        $permissionIssues = $stats_file | Where-Object {$_.Status -eq "ERROR_ACCESS_DENIED" -or $_.Status -eq "StatusAccessDenied"}
        if ($permissionIssues) {
            $permissionIssues | Format-Table -Autosize
        } else {
            Write-Host "No permission issues found" -ForegroundColor Green
        }
        
    } catch {
        Write-Error "Failed to analyze CompStatus data: $($_.Exception.Message)"
    }
}

# ========================
# Perfmon Collector Section
# ========================
function Setup-BHEPerfmon {
    param (
        [switch]$Delete
    )

    $collectorName = "BloodHound_System_Overview_Lite"
    $logPath = "C:\PerfLogs\$collectorName"
    $desktop = [Environment]::GetFolderPath("Desktop")
    $zipPath = Join-Path $desktop "${env:COMPUTERNAME}_PerfTrace.zip"

    if ($Delete) {
        Write-Host "Deleting collector $collectorName..."
        
        # Stop the collector if it's running
        $stopResult = & logman.exe stop $collectorName 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "Collector stopped successfully."
        } else {
            Write-Host "Collector was not running or already stopped."
        }
        
        # Wait a moment for the stop to complete
        Start-Sleep -Seconds 1
        
        # Delete the collector
        $deleteResult = & logman.exe delete $collectorName 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "Collector deleted successfully."
        } else {
            Write-Host "Failed to delete collector: $deleteResult" -ForegroundColor Yellow
        }
        
        # Verify deletion
        $null = & logman.exe query $collectorName 2>$null
        $deleted = ($LASTEXITCODE -ne 0)
        
        # Remove log directory
        if (Test-Path -LiteralPath $logPath) {
            try {
                Remove-Item -LiteralPath $logPath -Recurse -Force -ErrorAction SilentlyContinue
                Write-Host "Log directory removed."
            } catch {
                Write-Host "Could not remove log directory: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        
        if ($deleted) {
            Write-Host "Collector deletion completed successfully." -ForegroundColor Green
        } else {
            Write-Host "Collector may still exist (insufficient rights or in-use)." -ForegroundColor Yellow
        }
        return
    }

    # Ensure base log directory exists
    if (-not (Test-Path -LiteralPath (Split-Path $logPath -Parent))) {
        New-Item -ItemType Directory -Path (Split-Path $logPath -Parent) -Force | Out-Null
    }

    # Check if collector exists via exit code
    $null = & logman.exe query $collectorName 2>$null
    $collectorExists = ($LASTEXITCODE -eq 0)
    $didStop = $false

    if ($collectorExists) {
        $statusOut = (& logman.exe query $collectorName 2>$null) | Out-String
        if ($statusOut -match "Status\s*:\s*Running") {
            $choice = Read-Host "$collectorName is already running. Stop it? (Y to stop / Q to leave running and exit)"
            switch ($choice.ToUpper()) {
                "Y" { & logman.exe stop $collectorName 2>$null | Out-Null; Write-Host "Stopped collector."; $didStop = $true }
                "Q" { Write-Host "Leaving collector running."; return }
                default { Write-Host "Cancelled."; return }
            }
        } else {
            Write-Host "$collectorName exists but is not running. Starting it..."
            $startResult = & logman.exe start $collectorName 2>&1
            if ($LASTEXITCODE -ne 0) {
                Write-Host "Failed to start collector: $startResult" -ForegroundColor Red
                return
            }
            
            # Verify the collector actually started
            Start-Sleep -Seconds 2
            $verifyResult = & logman.exe query $collectorName 2>&1
            if ($LASTEXITCODE -eq 0) {
                $statusOut = $verifyResult | Out-String
                if ($statusOut -match "Status\s*:\s*Running") {
                    Write-Host "Collector $collectorName started and is running successfully." -ForegroundColor Green
                } else {
                    Write-Host "Collector failed to start properly. Status: $statusOut" -ForegroundColor Yellow
                }
            } else {
                Write-Host "Failed to verify collector status: $verifyResult" -ForegroundColor Red
            }
        }
    } else {
        Write-Host "Creating new collector $collectorName..."
        
        # Use logman command with all counters in one call (sample interval 30 seconds, max 512MB - can be changed here if needed)
        Write-Host "Creating collector with command: logman.exe create counter $collectorName -f bincirc -c `"\Process(*)\*`" `"\PhysicalDisk(*)\*`" `"\Processor(*)\*`" `"\Memory\*`" `"\Network Interface(*)\*`" `"\System\System Up Time`" -si 00:00:30 -max 512 -o $logPath -v mmddhhmm"
        $createResult = & logman.exe create counter $collectorName -f bincirc -c "\Process(*)\*" "\PhysicalDisk(*)\*" "\Processor(*)\*" "\Memory\*" "\Network Interface(*)\*" "\System\System Up Time" -si 00:00:30 -max 512 -o $logPath -v mmddhhmm 2>&1
        
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Failed to create collector: $createResult" -ForegroundColor Red
            return
        }
        
        Write-Host "Collector created successfully. Starting collection..."
        
        # Verify the collector was created with counters
        Write-Host "Verifying collector configuration..."
        $configResult = & logman.exe query $collectorName 2>&1
        if ($LASTEXITCODE -eq 0) {
            $configOut = $configResult | Out-String
            Write-Host "Collector configuration verified."
        } else {
            Write-Host "Warning: Could not verify collector configuration: $configResult" -ForegroundColor Yellow
        }
        
        # Start the collector
        $startResult = & logman.exe start $collectorName 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Failed to start collector: $startResult" -ForegroundColor Red
            return
        }
        
        # Verify the collector is actually running
        Start-Sleep -Seconds 2
        $verifyResult = & logman.exe query $collectorName 2>&1
        if ($LASTEXITCODE -eq 0) {
            $statusOut = $verifyResult | Out-String
            if ($statusOut -match "Status\s*:\s*Running") {
                Write-Host "Collector $collectorName started and is running successfully." -ForegroundColor Green
            } else {
                Write-Host "Collector created but failed to start properly. Status: $statusOut" -ForegroundColor Yellow
            }
        } else {
            Write-Host "Failed to verify collector status: $verifyResult" -ForegroundColor Red
        }
    }

    # Zip logs to Desktop only if we stopped the collector in this run
    if ($didStop) {
        try {
            Add-Type -AssemblyName 'System.IO.Compression.FileSystem' -ErrorAction SilentlyContinue
            if (Test-Path $zipPath) { Remove-Item $zipPath -Force -ErrorAction SilentlyContinue }
            $itemsToZip = @()
            
            # Add the log directory itself if it exists
            if (Test-Path -LiteralPath $logPath) { 
                $itemsToZip += $logPath 
                Write-Host "Added log directory to zip: $logPath"
            }
            
            # Look for .blg files that start with the collector name in the parent PerfLogs directory
            $parentLogPath = Split-Path $logPath -Parent  # This will be C:\PerfLogs
            $blgPattern = Join-Path $parentLogPath "$collectorName*.blg"
            Write-Host "Looking for .blg files with pattern: $blgPattern"
            
            $blgFiles = @(Get-ChildItem -Path $blgPattern -File -ErrorAction SilentlyContinue)
            
            if ($blgFiles -and $blgFiles.Count -gt 0) { 
                foreach ($file in $blgFiles) {
                    $itemsToZip += $file.FullName
                    Write-Host "Added .blg file to zip: $($file.Name)"
                }
                Write-Host "Found $($blgFiles.Count) performance log files to zip."
            } else {
                Write-Host "No .blg files found with pattern: $blgPattern" -ForegroundColor Yellow
                # Fallback: list all files in the parent PerfLogs directory to help debug
                $allFiles = @(Get-ChildItem -Path $parentLogPath -File -ErrorAction SilentlyContinue)
                if ($allFiles -and $allFiles.Count -gt 0) {
                    Write-Host "Files found in parent log directory ($parentLogPath):"
                    foreach ($file in $allFiles) {
                        Write-Host "  $($file.Name)"
                    }
                }
            }
            
            if ($itemsToZip.Count -gt 0) {
                Write-Host "Creating zip file..."
                $tmpZipFolder = Join-Path $env:TEMP ("bhe_perf_zip_" + [guid]::NewGuid())
                New-Item -ItemType Directory -Path $tmpZipFolder -Force | Out-Null
                
                foreach ($p in $itemsToZip) {
                    if (Test-Path -LiteralPath $p) {
                        Copy-Item -LiteralPath $p -Destination $tmpZipFolder -Recurse -Force -ErrorAction SilentlyContinue
                        Write-Host "Copied: $p"
                    } else {
                        Write-Host "Warning: Path not found: $p" -ForegroundColor Yellow
                    }
                }
                
                [System.IO.Compression.ZipFile]::CreateFromDirectory($tmpZipFolder, $zipPath)
                Remove-Item -LiteralPath $tmpZipFolder -Recurse -Force -ErrorAction SilentlyContinue
                Write-Host "Zipped perf logs to $zipPath" -ForegroundColor Green
            } else {
                Write-Host "No perf log files found to zip." -ForegroundColor Yellow
            }
        } catch {
            Write-Host "Failed to zip perf logs: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$workDir = Join-Path $OutputRoot "BHE_SupportLogs_$timestamp"
$zipPath = Join-Path $OutputRoot "BHE_SupportLogs_$timestamp.zip"

$transcriptPath = Join-Path $workDir 'collectorlogs.log'

Show-Banner

# Check if help was requested
if ($Help.IsPresent) {
    Show-CommandLineHelp
    Write-Host ""
    Write-Host "For more detailed information, see the README.md file." -ForegroundColor Cyan
    return
}

# Early handlers for perf-only switches
if ($GetBHEPerfmon.IsPresent) {
    Setup-BHEPerfmon
    return
}
if ($DeleteBHEPerfmon.IsPresent) {
    Setup-BHEPerfmon -Delete
    return
}

# Early handler for compstatus analysis
if ($PSBoundParameters.ContainsKey('GetCompStatus')) {
    Write-Host "CompStatus Analysis Mode" -ForegroundColor Cyan
    Write-Host ""
    
    if ([string]::IsNullOrWhiteSpace($GetCompStatus)) {
        Write-Error "GetCompStatus parameter requires a path to the compstatus.csv file."
        Write-Host "Usage: .\GetBHESupportLogsTool_Latest.ps1 -GetCompStatus <path_to_compstatus.csv>" -ForegroundColor Yellow
        return
    }
    
    if (-not (Test-Path -LiteralPath $GetCompStatus)) {
        Write-Error "CompStatus CSV file not found: $GetCompStatus"
        return
    }
    
    Write-Info "Analyzing compstatus.csv file: $GetCompStatus"
    
    # Perform the analysis
    Invoke-CompStatusAnalysis -CsvPath $GetCompStatus
    return
}

# Determine early whether this is a configuration-only invocation (no collection)
$__configOnlyParams = @('SetLogLevel', 'SetEnumerationLogLevel', 'RestartDelegator', 'SetAzureVerbosity', 'RestartAzureHound')
$__hasConfigParams = $false
foreach ($__p in $__configOnlyParams) {
    if ($PSBoundParameters.ContainsKey($__p)) { $__hasConfigParams = $true; break }
}
$__hasCollectionToggles = $PSBoundParameters.ContainsKey('ExcludeEventLogs') -or $PSBoundParameters.ContainsKey('ExcludeSettings') -or $PSBoundParameters.ContainsKey('LogArchiveNumber')
$__suppressCollectionWarnings = $__hasConfigParams -and -not $__hasCollectionToggles -and -not $All.IsPresent -and -not $Help.IsPresent

# Warn about potentially sensitive content (skip in configuration-only mode)
if (-not $__suppressCollectionWarnings) {
    Write-Warning 'This collection will include the below data!'
    if ($AllPlusPerf.IsPresent) { 
        Write-Host '-------> ALL logs will be collected: SharpHound, AzureHound, and Windows event logs.' -ForegroundColor DarkCyan
        Write-Host '-------> Performance Monitor trace will be started and run until manually stopped.' -ForegroundColor DarkCyan
    } elseif ($All.IsPresent) { 
        Write-Host '-------> ALL logs will be collected: SharpHound, AzureHound, and Windows event logs.' -ForegroundColor DarkCyan
        Write-Host '-------> Automated collection mode - no user input required.' -ForegroundColor DarkCyan
    } else {
        if (-not $ExcludeEventLogs) { Write-Host '-------> Windows Application and System event logs will be collected; use -ExcludeEventLogs to skip.' -ForegroundColor DarkCyan } 
        if (-not $ExcludeSettings) { Write-Host '-------> settings.json will be collected; use -ExcludeSettings to skip.' -ForegroundColor DarkCyan } 
    }
}

# Check if only configuration/service parameters are specified (no collection)
$configOnlyParams = @('SetLogLevel', 'SetEnumerationLogLevel', 'RestartDelegator', 'SetAzureVerbosity', 'RestartAzureHound')
$hasConfigParams = $false
foreach ($param in $configOnlyParams) {
    if ($PSBoundParameters.ContainsKey($param)) {
        $hasConfigParams = $true
        break
    }
}

# If only config params are specified, skip collection and just make changes
if ($hasConfigParams -and -not $PSBoundParameters.ContainsKey('ExcludeEventLogs') -and -not $PSBoundParameters.ContainsKey('ExcludeSettings') -and -not $PSBoundParameters.ContainsKey('LogArchiveNumber') -and -not $PSBoundParameters.ContainsKey('All') -and -not $PSBoundParameters.ContainsKey('Help')) {
    Write-Host "Configuration-only mode detected. Making requested changes..." -ForegroundColor Cyan
    
    # Status accumulator for changes
    $records = New-Object System.Collections.ArrayList
    
    # Handle SharpHound config changes and/or explicit restart request
    if ($PSBoundParameters.ContainsKey('SetLogLevel') -or $PSBoundParameters.ContainsKey('SetEnumerationLogLevel') -or $PSBoundParameters.ContainsKey('RestartDelegator')) {
        $svc = Get-ServiceObject -Name "SHDelegator"
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
        
        $desiredLevel = $SetLogLevel
        $desiredEnumLevel = $SetEnumerationLogLevel
        $didChange = $false
        if ($desiredLevel) { Set-BHELogLevel -ServiceProfilePath $profilePath -DesiredLevel $desiredLevel -StatusList $records; $didChange = $true }
        if ($desiredEnumLevel) { Set-BHEEnumerationLogLevel -ServiceProfilePath $profilePath -DesiredLevel $desiredEnumLevel -StatusList $records; $didChange = $true }
        if ($RestartDelegator) { Try-RestartDelegatorService -Name "SHDelegator" }
    }
    
    # Handle AzureHound config changes
    if ($PSBoundParameters.ContainsKey('SetAzureVerbosity')) { 
        Set-AzureHoundVerbosity -Verbosity $SetAzureVerbosity -StatusList $records 
    }
    if ($RestartAzureHound) { 
        Try-RestartAzureHoundService 
    }
    
    # Show summary of changes made
    Write-Host ""
    Write-Host "Configuration changes completed:" -ForegroundColor Green
    Print-Summary -StatusList $records -WorkDir "N/A" -ZipPath "N/A"
    return
}

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

$interactive = $Host.UI -and $Host.UI.RawUI
# Skip interactive prompt for -All and -AllPlusPerf to allow automated collection
if ($interactive -and -not $All.IsPresent -and -not $AllPlusPerf.IsPresent) {
    try {
        if (-not (Wait-ForEnter)) {
            Write-Host "Exiting by user request." -ForegroundColor Yellow
            return
        }
    } catch {
        Write-Host "Interactive mode failed, falling back to non-interactive..." -ForegroundColor Yellow
        $interactive = $false
    }
} elseif ($All.IsPresent) {
    Write-Host "Automated collection mode (-All): Proceeding without user input..." -ForegroundColor Green
    $interactive = $false
} elseif ($AllPlusPerf.IsPresent) {
    Write-Host "Automated collection mode (-AllPlusPerf): Proceeding without user input..." -ForegroundColor Green
    $interactive = $false
}

try {
    try {
        New-Item -ItemType Directory -Path $workDir -Force | Out-Null
    } catch {
        Write-Error "Failed to create output directory '$workDir': $($_.Exception.Message)"
        exit 1
    }

    try { 
        # Ensure no transcript is already running
        try { Stop-Transcript -ErrorAction SilentlyContinue | Out-Null } catch { }
        Start-Transcript -Path $transcriptPath -ErrorAction SilentlyContinue | Out-Null 
    } catch { }

    Write-Info "Output folder: $workDir"

    # Status accumulator used across collection stages
    $records = New-Object System.Collections.ArrayList

    # Determine collection mode
    $mode = 'SharpHound'
    if ($All.IsPresent -or $AllPlusPerf.IsPresent) {
        $mode = 'All'
    } elseif ($interactive) { 
        $mode = Prompt-SelectTarget 
    }

    # Resolve service and profile path (for settings.json and file collection) when SharpHound or All selected
    $svc = $null
    $profilePath = $null
    if ($mode -in @('SharpHound', 'All')) {
        $svc = Get-ServiceObject -Name "SHDelegator"
        if ($svc) {
            Write-Info "Using service '$($svc.Name)' (DisplayName: '$($svc.DisplayName)') running as '$($svc.StartName)'"
            $profilePath = Get-ProfilePathFromServiceStartName -StartName $svc.StartName
            if ($profilePath) {
                Write-Info "Resolved service profile path: $profilePath"
            } else {
                Write-Warn 'Could not resolve service profile path; will attempt current user profile.'
            }
        }
    }

    if ($mode -eq 'All') {
        Write-Host "Collecting ALL logs (SharpHound, AzureHound, and Event Logs)..." -ForegroundColor Cyan
        
        # SharpHound configuration changes
        $desiredLevel = $SetLogLevel
        $desiredEnumLevel = $SetEnumerationLogLevel
        $didChange = $false
        if ($desiredLevel) { Set-BHELogLevel -ServiceProfilePath $profilePath -DesiredLevel $desiredLevel -StatusList $records; $didChange = $true }
        if ($desiredEnumLevel) { Set-BHEEnumerationLogLevel -ServiceProfilePath $profilePath -DesiredLevel $desiredEnumLevel -StatusList $records; $didChange = $true }

        # AzureHound configuration changes
        if ($PSBoundParameters.ContainsKey('SetAzureVerbosity')) { Set-AzureHoundVerbosity -Verbosity $SetAzureVerbosity -StatusList $records }
        if ($RestartAzureHound) { Try-RestartAzureHoundService }

        # Event logs
        Collect-EventLogsWithStatus -DestinationFolder $workDir -StatusList $records
        
        # SharpHound files
        Collect-BHEFilesWithStatus -ServiceProfilePath $profilePath -DestinationFolder $workDir -StatusList $records
        
        # AzureHound files
        Collect-AzureHoundFilesWithStatus -DestinationFolder $workDir -StatusList $records

        # If AllPlusPerf, also ensure perfmon is set up and zip perf separately
        if ($AllPlusPerf.IsPresent) {
            try { Setup-BHEPerfmon } catch { Write-Warn "Perfmon setup failed: $($_.Exception.Message)" }
        }
        
    } elseif ($mode -eq 'SharpHound') {
        # Parameter-only: set BHE LogLevel and EnumerationLogLevel
        $desiredLevel = $SetLogLevel
        $desiredEnumLevel = $SetEnumerationLogLevel
        $didChange = $false
        if ($desiredLevel) { Set-BHELogLevel -ServiceProfilePath $profilePath -DesiredLevel $desiredLevel -StatusList $records; $didChange = $true }
        if ($desiredEnumLevel) { Set-BHEEnumerationLogLevel -ServiceProfilePath $profilePath -DesiredLevel $desiredEnumLevel -StatusList $records; $didChange = $true }

        # Event logs
        Collect-EventLogsWithStatus -DestinationFolder $workDir -StatusList $records
        # BHE files
        Collect-BHEFilesWithStatus -ServiceProfilePath $profilePath -DestinationFolder $workDir -StatusList $records
    } else {
        # AzureHound path
        if ($PSBoundParameters.ContainsKey('SetAzureVerbosity')) { Set-AzureHoundVerbosity -Verbosity $SetAzureVerbosity -StatusList $records }
        if ($RestartAzureHound) { Try-RestartAzureHoundService }
        # Event logs
        Collect-EventLogsWithStatus -DestinationFolder $workDir -StatusList $records
        # AzureHound files
        Collect-AzureHoundFilesWithStatus -DestinationFolder $workDir -StatusList $records
    }

    Write-Host ''
    try { Stop-Transcript | Out-Null } catch { }
    Write-Info 'Creating zip archive...'
    if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue }
    Compress-Archive -Path (Join-Path $workDir '*') -DestinationPath $zipPath -Force
    Add-Status -List $records -Name 'Zip Archive' -Path $zipPath -Status 'Created'
    Print-ItemResult -Name 'Zip Archive' -Status 'Created'

    Write-Host "Collection complete." -ForegroundColor Green
    Write-Host "Folder: $workDir"
    Write-Host "Zip:    $zipPath"

    Print-Summary -StatusList $records -WorkDir $workDir -ZipPath $zipPath
    if ($interactive) { Prompt-Open -WorkDir $workDir -ZipPath $zipPath }
} catch {
    $errorMsg = if ($_ -and $_.Exception) { $_.Exception.Message } else { "Unknown error occurred" }
    Write-Error "Unexpected error: $errorMsg"
} finally {
    try { Stop-Transcript | Out-Null } catch { }
}


