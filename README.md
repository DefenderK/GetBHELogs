
# BHE Logs Collector
Collect Windows event logs and BloodHound Enterprise (BHE) SharpHound service artifacts into a single zip.  
Designed for support and troubleshooting, with both interactive and fully non-interactive modes.

---

## Quick Start

From an **elevated PowerShell prompt** in the script directory:

```powershell
# Collect everything (default)
.\GetBHESupportLogs.ps1

# Collect logs but exclude event logs
.\GetBHESupportLogs.ps1 -ExcludeEventLogs

# Collect logs but exclude settings.json
.\GetBHESupportLogs.ps1 -ExcludeSettings

# Collect to a custom folder
.\GetBHESupportLogs.ps1 -OutputRoot "C:\Temp"

# Set log levels and restart service automatically
.\GetBHESupportLogs.ps1 -SetLogLevel Debug -SetEnumerationLogLevel Trace -RestartDelegatorAfterChange

# Run unattended (non-interactive/remote use)
.\GetBHESupportLogs.ps1 -AutoStart
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

---

## Requirements

* Windows **PowerShell 5.1+** (PowerShell 7+ also works).
* Recommended: Run as **Administrator** (for event log export and access to service profiles).
* Output folder (`-OutputRoot`) must exist and be writable. Defaults to the logged-on user’s Desktop.

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

```

### Interactive flow

* Displays an ASCII banner.
* Prompts: **Press Enter to collect logs, (H)elp for command line parameters, or Q to quit**.
* Displays per-item status as logs and files are collected.
* Prints a summary and offers: **Press O to open output folder, Z to open at zip, or any other key to exit**.

⚠️ **Note**: Log level changes and service restarts are controlled **only** via parameters (`-SetLogLevel`, `-SetEnumerationLogLevel`, `-RestartDelegatorAfterChange`).

---

## All parameters

* `-OutputRoot [string]` — Root folder where the output directory and zip are created. Defaults to Desktop.
* `-ServiceName [string]` — Delegator service name (default: `SHDelegator`).
* `-AutoStart [switch]` — Start collection immediately, suppressing prompts/dialogs (for remote/unattended runs).
* `-ExcludeEventLogs [switch]` — Skip exporting Windows Application/System event logs.
* `-ExcludeSettings [switch]` — Skip copying `settings.json` from the BHE folder.
* `-SetLogLevel [Trace|Debug|Information]` — Update LogLevel in `settings.json` before collection.
* `-SetEnumerationLogLevel [Trace|Debug|Information]` — Update EnumerationLogLevel in `settings.json`.
* `-RestartDelegatorAfterChange [switch]` — Automatically restart the Delegator service after log level changes.
* `-LogArchiveNumber [int]` — Copy only the N most recent files from the log\_archive folder.

---

## Notes

* Service resolution tries matches in this order:
  Name (`SHDelegator`), DisplayName (`SharpHoundDelegator`), Description (`SharpHound Delegation Service`).
* If not found, BHE file collection may show as *NotFound*.
* **Privacy:** Event logs may contain PII; `settings.json` may contain endpoints or config.
  Use `-ExcludeEventLogs` and/or `-ExcludeSettings` if needed.

---

## Output example

* Folder: `BHE_SupportLogs_YYYYMMDD_HHMMSS`
* Zip: `BHE_SupportLogs_YYYYMMDD_HHMMSS.zip`
* Transcript: `collectorlogs.log` inside the folder

---

## Troubleshooting

* If EVTX export fails, the script falls back to XML export via `Get-WinEvent`.
* If BHE files are *NotFound*, ensure the `SHDelegator` service is installed and running, and that your account has permissions to access the service profile.

---

## Screenshots / Demo Output

### Startup

```
========================================
        BHE Logs Collector v.1.1        
========================================
WARNING: Note: This collection will include the below data!
WARNING: Windows Application and System event logs will be collected; use -ExcludeEventLogs to skip.
WARNING: settings.json will be collected; use -ExcludeSettings to skip.
Press Enter to collect logs, (H)elp for parameters, or Q to quit
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
\nCollection complete.
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
Open folder: file:///C:/Users/administrator.DEFENDERK/Desktop/BHE_SupportLogs_20250821_092237
Open zip:    file:///C:/Users/administrator.DEFENDERK/Desktop/BHE_SupportLogs_20250821_092237.zip

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




