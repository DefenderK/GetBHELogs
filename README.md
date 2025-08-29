
# BHE Logs Collector
Collector tool to gather: Windows event logs, BloodHound Enterprise (BHE) SharpHound and/or AzureHound service artifacts, and/or a Performance monitor trace into an output zip.  
Designed for support and troubleshooting.


---

## What it does

* Exports **Application** and **System** event logs (`.evtx`; falls back to XML if needed).
* Collects BHE artifacts from the **SharpHound service account profile**:
  * `BloodHoundEnterprise/log_archive/`
  * `BloodHoundEnterprise/service.log`
  * `BloodHoundEnterprise/settings.json`
* When AzureHound is selected, in addition to the event logs, collects `azurehound.log` from `C:\Program Files\AzureHound Enterprise\azurehound.log` if present.
* Shows per-item status and a final summary.
* Creates a **timestamped folder and zip** in the chosen output directory (Desktop by default).
* When using `-All`, collects **all** logs simultaneously: SharpHound, AzureHound, and Windows event logs.
* When using `-AllPlusPerf`, it additionally creates a Performance Monitor Data Collector Set and starts the trace. It creates the output blg file in `C:\PerfLogs`

---

## Requirements

* Windows **PowerShell 5.1+** (PowerShell 7+ also works).
* Recommended: Run as **Administrator** (for event log export and access to service profiles).
* Output folder (`-OutputRoot`) must exist and be writable. Defaults to the logged-on user's Desktop.

---

## Quick Start

1. **Download** the script to your target system (or git clone)
2. **Open PowerShell as Administrator**
3. **Navigate** to the script directory
4. **Run** the script:
   ```powershell
   .\GetBHESupportLogsTool.ps1
   ```
5. **Follow prompts** to collect logs
6. **Review** the generated zip file and folder

For automated collection:
```powershell
.\GetBHESupportLogsTool.ps1 -All
```

---

## Usage

From an elevated PowerShell prompt in the scripts directory:

```powershell
# Run the script directly (defaults to interactive mode)
.\GetBHESupportLogsTool.ps1

```

### Interactive flow

* Displays an ASCII banner.
* Prompts: **Press Enter to collect logs, or Q to quit**.
* Displays output log location.
* Prompts for: Select collection target: (S)harpHound or (A)zureHound
  Choice [S/A]:
* Displays per-item status as logs and files are collected.
* Prints a summary and offers: **Press O to open output folder, Z to open at zip, or any other key to exit**.
* When using `-All`, all logs are collected regardless of interactive target selection.

**Note**: The script is interactive by default for selective collection. Use `-All` or `-AllPlusPerf` to collect everything automatically without user input, or run without parameters for selective collection.

### Examples

#### Basic Collection
```powershell
# Interactive collection (default)
.\GetBHESupportLogsTool.ps1

# Automated collection of all logs
.\GetBHESupportLogsTool.ps1 -All

# Collection with custom output location
.\GetBHESupportLogsTool.ps1 -OutputRoot "C:\Temp"
```

#### Configuration Management
```powershell
# Set SharpHound logging levels and restart service
.\GetBHESupportLogsTool.ps1 -SetLogLevel Debug -SetEnumerationLogLevel Trace -RestartDelegator

# Set AzureHound verbosity and restart
.\GetBHESupportLogsTool.ps1 -SetAzureVerbosity 2 -RestartAzureHound
```

#### Selective Collection
```powershell
# Skip Event Logs and settings.json
.\GetBHESupportLogsTool.ps1 -ExcludeEventLogs -ExcludeSettings

# Limit log archive collection
.\GetBHESupportLogsTool.ps1 -LogArchiveNumber 10

# Performance monitoring only
.\GetBHESupportLogsTool.ps1 -GetBHEPerfmon
```

#### Help & Information
```powershell
# Display help
.\GetBHESupportLogsTool.ps1 -Help
```

### Configuration-only mode
* When using **only** configuration/service parameters (`-SetLogLevel`, `-SetEnumerationLogLevel`, `-RestartDelegator`, `-SetAzureVerbosity`, `-RestartAzureHound`), the script skips the collection options entirely.
* Only makes the requested changes and shows verification of what was updated.
* Useful for troubleshooting when you need to change settings but don't want to collect logs yet.
* Example: `.\GetBHESupportLogsTool.ps1 -SetAzureVerbosity 2 -RestartAzureHound` will only change verbosity and restart the service.

⚠️ **Note**: Log level changes and service restarts are controlled **only** via parameters.

---

## All Parameters

### Collection Control
* `-OutputRoot [string]` — Root folder where the output directory and zip are created. Defaults to Desktop.
* `-All [switch]` — Collect all logs: SharpHound, AzureHound, and Windows event logs simultaneously. **Automated execution - no user input required.**
* `-AllPlusPerf [switch]` — Do everything `-All` does and also ensure a BHE perfmon trace is set up. **Automated execution - no user input required.**
* `-LogArchiveNumber [int]` — Copy only the N most recent files from the log_archive folder.

### Exclusion Options
* `-ExcludeEventLogs [switch]` — Skip exporting Windows Application/System event logs.
* `-ExcludeSettings [switch]` — Skip copying `settings.json` from the BHE folder.

### SharpHound Configuration Management
* `-SetLogLevel [Trace|Debug|Information]` — Update LogLevel in `settings.json` before collection.
* `-SetEnumerationLogLevel [Trace|Debug|Information]` — Update EnumerationLogLevel in `settings.json`.
* `-RestartDelegator [switch]` — Automatically restart the Delegator service (useful after log level changes).

### AzureHound Configuration Management
* `-SetAzureVerbosity [0|1|2]` — Set AzureHound service log verbosity in `C:\ProgramData\azurehound\config.json` (0=Default, 1=Debug, 2=Trace).
* `-RestartAzureHound [switch]` — Restart the `AzureHound` Windows service (useful after log level changes).

### Performance Monitoring
* `-GetBHEPerfmon [switch]` — Perfmon-only mode. If the Data Collector Set is running, you'll be prompted to stop it and then the trace files in `C:\PerfLogs` are zipped to Desktop as `<COMPUTERNAME>_PerfTrace.zip`. If it isn't present, the Data Collector Set is created and started with recommended counters.
* `-DeleteBHEPerfmon [switch]` — Stop and delete the Data Collector Set.

### Utility
* `-Help [switch]` — Display command line parameters and examples, then exit.


---

## Performance Monitor tracing

The script can manage a lightweight performance monitor trace using Windows `logman`:

- Data Collector Set name: `BloodHound_System_Overview_Lite`
- Location: `C:\PerfLogs`
- Format: binary circular log (`bincirc`), 512 MB max, 30s sample interval
- Counters included: `"\Process(*)\*" "\PhysicalDisk(*)\*" "\Processor(*)\*" "\Memory\*" "\Network Interface(*)\*" "\System\System Up Time"`
- Note: You can also run `logman query` to check if the Data Collector Set is already setup and trace is running, example output below:
  ```
  PS C:\Users\administrator.DEFENDERK\Desktop> logman query

  Data Collector Set                      Type                          Status
  -------------------------------------------------------------------------------
  BloodHound_System_Overview_Lite         Counter                       Running 
  ```

### Typical flows

- Start or check the Data Collector Set, and if the trace is already running choose to stop and zip:
  ```powershell
  .\GetBHESupportLogsTool.ps1 -GetBHEPerfmon
  # If running: press Y to stop and zip to Desktop as <COMPUTERNAME>_PerfTrace.zip
  # Press Q to leave it running; any other key cancels
  ```

- Collect all logs and also ensure the Data Collector Set is set up (automated execution, does not stop/zip automatically):
  ```powershell
  .\GetBHESupportLogsTool.ps1 -AllPlusPerf
  # Runs automatically without user input
  # Later, run -GetBHEPerfmon and choose Y to stop and zip
  ```

- Delete the Data Collector Set:
  ```powershell
  .\GetBHESupportLogsTool.ps1 -DeleteBHEPerfmon
  ```

---

## Notes

* **Privacy:** Event logs may contain PII; `settings.json` may contain endpoints or config.
  Use `-ExcludeEventLogs` and/or `-ExcludeSettings` if needed.

---

## Output example

* Folder: `BHE_SupportLogs_YYYYMMDD_HHMMSS`
* Zip: `BHE_SupportLogs_YYYYMMDD_HHMMSS.zip`
* Perf Zip: `<COMPUTERNAME>_PerfTrace.zip`
* Tool Collector Transcript: `collectorlogs.log` inside the folder

---

## Troubleshooting

* If EVTX export fails, the script falls back to XML export via `Get-WinEvent`.
* If BHE files are *NotFound*, ensure the `SHDelegator` or `AzureHound` service is installed and running, and that your account has permissions to access the service profile.

---


## Demo Output Example

### Startup (Standard Mode)

```
========================================
        BHE Logs Collector v.2.9       
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

---

## License

This project is licensed under the **MIT License**.
You are free to use, modify, and distribute it with attribution. See the [LICENSE](LICENSE) file for details.

```

---


```

















