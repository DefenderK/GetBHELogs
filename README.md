````markdown
# BHE Logs Collector

![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue?logo=powershell)
![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)
![Last Commit](https://img.shields.io/github/last-commit/your-org-or-user/your-repo)

Collect Windows event logs and BloodHound Enterprise (BHE) service artifacts into a single zip.  
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
===========================================
   BHE Logs Collector
===========================================
Press Enter to collect logs, (H)elp for command line parameters, or Q to quit
```

### Collection Progress

```
[OK] Application event log exported (1200 entries)
[OK] System event log exported (850 entries)
[OK] service.log copied
[OK] settings.json copied
[OK] 10 files from log_archive copied
```

### Summary

```
Collection complete.
Output folder: C:\Users\Alice\Desktop\BHE_SupportLogs_20250120_143215
Zip created:   C:\Users\Alice\Desktop\BHE_SupportLogs_20250120_143215.zip

Press O to open output folder, Z to open at zip, or any other key to exit.
```

---

## License

This project is licensed under the **MIT License**.
You are free to use, modify, and distribute it with attribution. See the [LICENSE](LICENSE) file for details.

```

---


```
