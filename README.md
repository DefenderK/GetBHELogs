
# BHE Logs Collector 2.0
Collector tool to get: Windows event logs, BloodHound Enterprise (BHE) SharpHound and/or AzureHound service artifacts into a single zip.  
Designed for support and troubleshooting.

---

## Quick Start

From an **elevated PowerShell prompt** in the script directory:

```powershell
# Default collection
.\GetBHESupportLogs.ps1

# Collect all logs (SharpHound + AzureHound + Event Logs) in one run
.\GetBHESupportLogs.ps1 -All

# Show command line help
.\GetBHESupportLogs.ps1 -Help

# Collect logs but exclude event logs
.\GetBHESupportLogs.ps1 -ExcludeEventLogs

# Collect logs but exclude settings.json
.\GetBHESupportLogs.ps1 -ExcludeSettings

# Collect to a custom folder (default is Desktop)
.\GetBHESupportLogs.ps1 -OutputRoot "C:\Temp"

# Set log levels for SharpHound and restart service automatically
.\GetBHESupportLogs.ps1 -SetLogLevel Debug -SetEnumerationLogLevel Trace -RestartDelegator


````

---

## What it does

* Exports **Application** and **System** event logs (`.evtx`; falls back to XML if needed).
* Collects BHE artifacts from the **Delegator service account profile**:
  * `BloodHoundEnterprise/log_archive/`
  * `BloodHoundEnterprise/service.log`
  * `BloodHoundEnterprise/settings.json`

* Creates a **timestamped folder and zip** in the chosen output directory (Desktop by default).
* Shows per-item status and a final summary.
* When AzureHound is selected, in addition to the event logs, collects `azurehound.log` from `C:\Program Files\AzureHound Enterprise\azurehound.log` if present.
* When using `-All`, collects **all** logs simultaneously: SharpHound, AzureHound, and Windows event logs.

---

## Requirements

* Windows **PowerShell 5.1+** (PowerShell 7+ also works).
* Recommended: Run as **Administrator** (for event log export and access to service profiles).
* Output folder (`-OutputRoot`) must exist and be writable. Defaults to the logged-on user's Desktop.

---

## Usage

From an elevated PowerShell prompt in the script directory:

```powershell
# Run the script directly (defaults to interactive mode)
.\GetBHESupportLogs.ps1
```

### Optional parameter examples

```powershell
# Choose a different output folder
.\GetBHESupportLogs.ps1 -OutputRoot "C:\Temp"

# Skip event logs or settings.json if customers are sensitive to their contents
.\GetBHESupportLogs.ps1 -ExcludeEventLogs
.\GetBHESupportLogs.ps1 -ExcludeSettings

# Set BHE logging levels (validated: Trace, Debug, Information)
.\GetBHESupportLogs.ps1 -SetLogLevel Debug
.\GetBHESupportLogs.ps1 -SetEnumerationLogLevel Trace

# Limit log_archive collection to N most recent files (default is to collect all)
.\GetBHESupportLogs.ps1 -LogArchiveNumber 10

# AzureHound: set verbosity to Trace (2) and restart the service
.\GetBHESupportLogs.ps1 -SetAzureVerbosity 2 -RestartAzureHound

# AzureHound: revert to default verbosity (0) without restart
.\GetBHESupportLogs.ps1 -SetAzureVerbosity 0

# Collect all logs with custom output location
.\GetBHESupportLogs.ps1 -All -OutputRoot "C:\Temp"

```

### Interactive flow

* Displays an ASCII banner.
* Prompts: **Press Enter to collect logs, or Q to quit**.
* Displays ouput log location.
* Prompts: Select collection target: (S)harpHound or (A)zureHound
  Choice [S/A]:
* Displays per-item status as logs and files are collected.
* Prints a summary and offers: **Press O to open output folder, Z to open at zip, or any other key to exit**.
* When using `-All`, all logs are collected regardless of interactive target selection.

**Note**: The script is always interactive by default. Use `-All` to collect everything automatically, or run without parameters for selective collection.

### Configuration-only mode

* When using **only** configuration/service parameters (`-SetLogLevel`, `-SetEnumerationLogLevel`, `-RestartDelegatorAfterChange`, `-SetAzureVerbosity`, `-RestartAzureHound`), the script skips the collection options entirely.
* Only makes the requested changes and shows verification of what was updated.
* Useful for troubleshooting when you need to change settings but don't want to collect logs yet.
* Example: `.\GetBHESupportLogsTool.ps1 -SetAzureVerbosity 2 -RestartAzureHound` will only change verbosity and restart the service.

⚠️ **Note**: Log level changes and service restarts are controlled **only** via parameters (`-SetLogLevel`, `-SetEnumerationLogLevel`, `-RestartDelegatorAfterChange`).

---

## All parameters

* `-OutputRoot [string]` — Root folder where the output directory and zip are created. Defaults to Desktop.
* `-All [switch]` — Collect all logs: SharpHound, AzureHound, and Windows event logs simultaneously.
* `-Help [switch]` — Display command line parameters and examples, then exit.
* `-ExcludeEventLogs [switch]` — Skip exporting Windows Application/System event logs.
* `-ExcludeSettings [switch]` — Skip copying `settings.json` from the BHE folder.
* `-SetLogLevel [Trace|Debug|Information]` — Update LogLevel in `settings.json` before collection.
* `-SetEnumerationLogLevel [Trace|Debug|Information]` — Update EnumerationLogLevel in `settings.json`.
* `-RestartDelegator [switch]` — Automatically restart the Delegator service (useful after log level changes).
* `-LogArchiveNumber [int]` — Copy only the N most recent files from the log_archive folder.
* `-SetAzureVerbosity [0|1|2]` — Set AzureHound service log verbosity in `C:\ProgramData\azurehound\config.json` (0=Default, 1=Debug, 2=Trace).
* `-RestartAzureHound [switch]` — Restart the `AzureHound` Windows service.

---

## Notes

* Service resolution tries matching in this order:
  Name (`SHDelegator`), DisplayName (`SharpHoundDelegator`), Description (`SharpHound Delegation Service`).
* If not found, BHE file collection may show as *NotFound*.
* AzureHound service name and display name: `AzureHound`. Description: "The official tool for collecting Azure data for BloodHound and BloodHound Enterprise." Configuration file path: `C:\ProgramData\azurehound\config.json` (`verbosity` set to 0, 1, or 2).
* **Privacy:** Event logs may contain PII; `settings.json` may contain endpoints or config.
  Use `-ExcludeEventLogs` and/or `-ExcludeSettings` if needed.

---

## Output example

* Folder: `BHE_SupportLogs_YYYYMMDD_HHMMSS`
* Zip: `BHE_SupportLogs_YYYYMMDD_HHMMSS.zip`
* Tool Collector Transcript: `collectorlogs.log` inside the folder

---

## Troubleshooting

* If EVTX export fails, the script falls back to XML export via `Get-WinEvent`.
* If BHE files are *NotFound*, ensure the `SHDelegator` or `AzureHound` service is installed and running, and that your account has permissions to access the service profile.

---

## The `-All` Parameter

The `-All` parameter is a quick and easy feature that allows you to collect **all** logs in a single run:

* **SharpHound logs**: service.log, settings.json, log_archive folder contents
* **AzureHound logs**: azurehound.log  
* **Windows Event Logs**: Application and System logs

### Examples:

```powershell
# Collect everything in one command
.\GetBHESupportLogs.ps1 -All

# Collect all logs to a specific location
.\GetBHESupportLogs.ps1 -All -OutputRoot "C:\Logs"

### Benefits:

* **Single command**: No need to run multiple collection cycles
* **Consistent output**: All logs are collected at the same timestamp
* **Complete coverage**: Ensures nothing is missed during collection

```

## Screenshots / Demo Output Example

### Startup

```
========================================
        BHE Logs Collector v.2.0       
========================================
WARNING: This collection will include the below data!
-------> Windows Application and System event logs will be collected; use -ExcludeEventLogs to skip.
-------> settings.json will be collected; use -ExcludeSettings to skip.
Press Enter to collect logs, or Q to quit
```

### Collection Progress

```
[INFO] Output folder: C:\Users\AdminUser\Desktop\BHE_SupportLogs_20250821_092237
[INFO] Using service 'SHDelegator' (DisplayName: 'SharpHoundDelegator') running as 'DOMAIN\svc_sharphound'
[INFO] Resolved service profile path: C:\Users\svc_sharphound
Collecting Windows Event Logs...
[INFO] Exporting Application and System event logs...
  - Application Event Log ... Collected - EVTX
  - System Event Log ... Collected - EVTX
Collecting BloodHoundEnterprise files...
  - BHE log_archive ... Collected
  - BHE service.log ... Collected
  - BHE settings.json ... Collected

[INFO] Creating zip archive...
  - Zip Archive ... Created
Collection complete.
Folder: C:\Users\AdminUser\Desktop\BHE_SupportLogs_20250821_092237
Zip:    C:\Users\AdminUser\Desktop\BHE_SupportLogs_20250821_092237.zip
```

### Summary

```
Collected:
  - Application Event Log -> C:\Users\AdminUser\Desktop\BHE_SupportLogs_20250821_092237\Application.evtx
  - System Event Log -> C:\Users\AdminUser\Desktop\BHE_SupportLogs_20250821_092237\System.evtx
  - BHE log_archive -> C:\Users\AdminUser\Desktop\BHE_SupportLogs_20250821_092237\BloodHoundEnterprise\log_archive
  - BHE service.log -> C:\Users\AdminUser\Desktop\BHE_SupportLogs_20250821_092237\BloodHoundEnterprise\service.log
  - BHE settings.json -> C:\Users\AdminUser\Desktop\BHE_SupportLogs_20250821_092237\BloodHoundEnterprise\settings.json
  - Zip Archive -> C:\Users\AdminUser\Desktop\BHE_SupportLogs_20250821_092237.zip

Output folder: C:\Users\AdminUser\Desktop\BHE_SupportLogs_20250821_092237
Zip archive:  C:\Users\AdminUser\Desktop\BHE_SupportLogs_20250821_092237.zip

Press O to open output folder, Z to open at zip, or any other key to exit.
Choice: 
```

---

## License

This project is licensed under the **MIT License**.
You are free to use, modify, and distribute it with attribution. See the [LICENSE](LICENSE) file for details.

```

---


```








