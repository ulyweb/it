<!-- Badges -->
![Version](https://img.shields.io/badge/Version-1.0-blue?style=flat-square)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?style=flat-square&logo=powershell)
![Platform](https://img.shields.io/badge/Platform-Windows%2011-0078D6?style=flat-square&logo=windows)
![WPF](https://img.shields.io/badge/UI-WPF%20GUI-purple?style=flat-square)
![WinRM](https://img.shields.io/badge/Transport-WinRM-orange?style=flat-square)
![License](https://img.shields.io/badge/License-Internal%20Use-red?style=flat-square)

---

# IT_RemoteDeploy

**A WPF-based PowerShell GUI tool for secure, auditable file transfers between a local IT workstation and remote Windows endpoints over WinRM.**

IT_RemoteDeploy provides IT support technicians with a visual, point-and-click interface for pushing scripts, configs, or diagnostic tools to remote machines -- or pulling logs, reports, and captured data back to the local workstation. Every transfer is SHA-256 verified, logged in the console, and designed for enterprise auditability.

Built as a companion tool to [IT_HealthCheck](./README.md), IT_RemoteDeploy closes the deployment loop: **push the diagnostic tool out, run it remotely, then pull the results back.**

---

## Table of Contents

1. [Features](#features)
2. [Requirements](#requirements)
3. [Quick Start](#quick-start)
4. [Architecture Overview](#architecture-overview)
5. [Connection Test (3-Stage Validation)](#connection-test-3-stage-validation)
6. [Transfer Modes](#transfer-modes)
   - [Push Mode (Local to Remote)](#push-mode-local-to-remote)
   - [Pull Mode (Remote to Local)](#pull-mode-remote-to-local)
7. [Pull Mode Wildcard Examples](#pull-mode-wildcard-examples)
8. [Step-by-Step Usage Walkthrough](#step-by-step-usage-walkthrough)
9. [Security and Access Control](#security-and-access-control)
10. [Error Handling](#error-handling)
11. [Integration with IT_HealthCheck](#integration-with-it_healthcheck)
12. [Design Principles](#design-principles)
13. [Pro Tips](#pro-tips)
14. [File Structure](#file-structure)
15. [License](#license)

---

## Features

| Category | Capability |
|---|---|
| **GUI Interface** | Modern WPF window with dark theme, real-time status updates, and scrollable console log |
| **Push Mode** | Transfer a single file from the local IT workstation to a remote endpoint |
| **Pull Mode** | Transfer one or more files (with wildcard support) from a remote endpoint to the local workstation |
| **Connection Test** | 3-stage pre-flight validation: Ping, WinRM session, and Admin group membership check |
| **SHA-256 Verification** | Every file transfer is integrity-verified using SHA-256 hash comparison (source vs destination) |
| **Wildcard Pull** | Pull multiple files from a remote directory using patterns like `*.csv`, `*_capture_log_*.csv`, etc. |
| **Background Execution** | All network operations run in a PowerShell background runspace to keep the UI responsive |
| **Audit Logging** | Every action, result, and error is timestamped and logged in the embedded console panel |
| **Non-Destructive** | No files are overwritten without user-initiated action; no destructive changes to remote systems |

---

## Requirements

| Requirement | Details |
|---|---|
| **Operating System** | Windows 10 / Windows 11 |
| **PowerShell** | Version 5.1 or later |
| **WinRM** | Enabled on the remote target machine (`winrm quickconfig`) |
| **Network** | TCP port 5985 (HTTP) or 5986 (HTTPS) open between local and remote |
| **Permissions** | The executing account must be a member of the local Administrators group on the remote machine |
| **Run As** | Must be launched as Administrator on the local IT workstation |

---

## Quick Start

```powershell
# 1. Open PowerShell as Administrator
# 2. Navigate to the script directory
cd C:\IT_Scripts

# 3. Run the tool
powershell.exe -ExecutionPolicy Bypass -File "IT_RemoteDeploy.ps1"
```

The WPF GUI window will launch. Enter the remote machine name or IP, then choose Push or Pull mode.

---

## Architecture Overview

```
+-----------------------------------------------------+
|              IT_RemoteDeploy (WPF GUI)               |
|-----------------------------------------------------|
|                                                     |
|  +-------------+   +----------------------------+   |
|  | Remote Host |   |  Transfer Mode Selector    |   |
|  | Input Field |   |  [Push]  or  [Pull]        |   |
|  +-------------+   +----------------------------+   |
|                                                     |
|  +-----------------------------------------------+  |
|  |         Transfer Configuration Panel          |  |
|  |  - Local Path       - Remote Path             |  |
|  |  - File Browser     - Wildcard Pattern        |  |
|  +-----------------------------------------------+  |
|                                                     |
|  +-----------------------------------------------+  |
|  |         Action Buttons                        |  |
|  |  [Test Connection]  [Execute Transfer]        |  |
|  +-----------------------------------------------+  |
|                                                     |
|  +-----------------------------------------------+  |
|  |         Console / Audit Log (ScrollViewer)    |  |
|  |  [timestamp] Connection test passed...        |  |
|  |  [timestamp] Pushing file to remote...        |  |
|  |  [timestamp] SHA-256 verified. Transfer OK.   |  |
|  +-----------------------------------------------+  |
+-----------------------------------------------------+
        |                           |
        | WinRM (PS Remoting)       | Local File I/O
        v                           v
+----------------+          +----------------+
| Remote Machine |          | Local IT       |
| (Target)       |          | Workstation    |
+----------------+          +----------------+
```

**Key Design Points:**

- The WPF UI runs on the main thread.
- All network-bound operations (connection test, file copy, hash verification) execute in a **PowerShell background runspace** to prevent the GUI from freezing.
- Status updates are marshaled back to the UI thread via the WPF Dispatcher.

---

## Connection Test (3-Stage Validation)

Before any file transfer is attempted, the tool runs a 3-stage pre-flight check:

| Stage | Test | What It Does | Pass Criteria |
|---|---|---|---|
| **1** | ICMP Ping | Sends a ping to the remote hostname or IP | Remote machine responds to ping |
| **2** | WinRM Session | Attempts to create a `New-PSSession` to the remote machine | Session established successfully |
| **3** | Admin Verify | Checks if the current user is in the local Administrators group on the remote machine | User is confirmed as local admin |

- If **any stage fails**, the transfer buttons are disabled and the console displays a clear error message with remediation guidance.
- If **all 3 stages pass**, the console shows a green confirmation and the transfer is unlocked.

---

## Transfer Modes

### Push Mode (Local to Remote)

> Transfer a file **from your local IT workstation to the remote endpoint**.

| Field | Description |
|---|---|
| **Local File Path** | Full path to the file on your local machine (e.g., `C:\IT_Scripts\IT_HealthCheck.ps1`) |
| **Remote Destination** | Directory on the remote machine where the file will be placed (e.g., `C:\IT_Scripts\`) |

**Push Workflow:**

1. Select "Push" mode in the UI.
2. Browse or type the local file path.
3. Enter the remote destination directory.
4. Click **Execute Transfer**.
5. The tool copies the file via `Copy-Item -ToSession`.
6. SHA-256 hash is computed on both the local source and remote destination.
7. Hashes are compared -- if they match, the console reports **Transfer Verified**.
8. If hashes do not match, the console reports a **Hash Mismatch Warning**.

---

### Pull Mode (Remote to Local)

> Transfer one or more files **from the remote endpoint to your local IT workstation**.

| Field | Description |
|---|---|
| **Remote Source Path** | Full path or wildcard pattern on the remote machine (e.g., `C:\Temp\*_capture_log_*.csv`) |
| **Local Destination** | Directory on your local machine where files will be saved (e.g., `C:\Temp\`) |

**Pull Workflow:**

1. Select "Pull" mode in the UI.
2. Enter the remote source path (supports wildcards).
3. Enter or browse the local destination directory.
4. Click **Execute Transfer**.
5. The tool resolves wildcard patterns on the remote machine to get the list of matching files.
6. Each file is copied via `Copy-Item -FromSession`.
7. SHA-256 hash is computed **per file** on both the remote source and local destination.
8. Results are logged individually in the console for full auditability.

---

## Pull Mode Wildcard Examples

| Pattern | What It Matches |
|---|---|
| `C:\Temp\*.csv` | All CSV files in the remote `C:\Temp` directory |
| `C:\Temp\*_capture_log_*.csv` | All IT_HealthCheck capture log CSVs |
| `C:\Temp\*.html` | All HTML dashboard reports |
| `C:\Temp\HOSTNAME_*.html` | HTML reports for a specific device |
| `C:\Temp\HOSTNAME_capture_log_202505*.csv` | Capture logs from May 2025 for a specific device |
| `C:\IT_Scripts\*.ps1` | All PowerShell scripts in the remote IT_Scripts folder |

> **Note:** The wildcard is resolved on the remote machine using `Get-ChildItem`. If no files match, the console will report "No files matched the specified pattern."

---

## Step-by-Step Usage Walkthrough

### Scenario A: Push IT_HealthCheck to a Remote Machine

```
1. Launch IT_RemoteDeploy as Administrator.
2. Enter remote machine name:  PC-EXEC-0042
3. Click [Test Connection].
4. Wait for all 3 stages to pass (Ping -> WinRM -> Admin).
5. Select mode: Push
6. Local File Path:      C:\IT_Scripts\IT_HealthCheck.ps1
7. Remote Destination:   C:\IT_Scripts\
8. Click [Execute Transfer].
9. Console output:
   [12:30:01] Pushing IT_HealthCheck.ps1 to PC-EXEC-0042...
   [12:30:04] File copied successfully.
   [12:30:05] Local  SHA-256: a1b2c3d4e5f6...
   [12:30:06] Remote SHA-256: a1b2c3d4e5f6...
   [12:30:06] SHA-256 MATCH -- Transfer Verified.
```

### Scenario B: Pull Capture Logs from a Remote Machine

```
1. Launch IT_RemoteDeploy as Administrator.
2. Enter remote machine name:  PC-EXEC-0042
3. Click [Test Connection].
4. Wait for all 3 stages to pass.
5. Select mode: Pull
6. Remote Source Path:   C:\Temp\*_capture_log_*.csv
7. Local Destination:    C:\Temp\
8. Click [Execute Transfer].
9. Console output:
   [12:35:01] Resolving wildcard on PC-EXEC-0042...
   [12:35:02] Found 3 matching files.
   [12:35:02] Pulling PC-EXEC-0042_capture_log_20260512_143022.csv...
   [12:35:04] SHA-256 MATCH -- Verified.
   [12:35:04] Pulling PC-EXEC-0042_capture_log_20260511_091505.csv...
   [12:35:06] SHA-256 MATCH -- Verified.
   [12:35:06] Pulling PC-EXEC-0042_capture_log_20260510_160830.csv...
   [12:35:08] SHA-256 MATCH -- Verified.
   [12:35:08] All 3 files transferred and verified successfully.
```

---

## Security and Access Control

| Control | Details |
|---|---|
| **Authentication** | Uses the current Windows identity. No credentials are stored or prompted. |
| **Authorization** | The executing user must be in the local Administrators group on the remote machine. |
| **WinRM Transport** | Uses PowerShell Remoting (`New-PSSession`) over WinRM (default: HTTP on port 5985). |
| **Integrity Verification** | Every file is SHA-256 hashed at source and destination. Mismatches are flagged immediately. |
| **Audit Trail** | All actions are timestamped in the console log. No silent failures. |
| **No Credential Storage** | The tool never saves, caches, or exports credentials. |
| **No Destructive Actions** | The tool only copies files. It does not delete, modify, or execute files on the remote machine. |

### WinRM Prerequisites on the Remote Machine

```powershell
# Run on the remote machine as Administrator:
winrm quickconfig

# Or enable PS Remoting directly:
Enable-PSRemoting -Force
```

---

## Error Handling

| Error Scenario | How It Is Handled |
|---|---|
| Remote machine unreachable (Ping fails) | Console displays "Stage 1 FAILED: Host unreachable." Transfer buttons disabled. |
| WinRM not enabled on remote | Console displays "Stage 2 FAILED: Cannot establish WinRM session." Guidance displayed. |
| User not in Administrators group | Console displays "Stage 3 FAILED: Access denied. Current user is not a remote admin." |
| Local file not found (Push mode) | Console displays "ERROR: Local file not found at specified path." |
| No matching files (Pull wildcard) | Console displays "No files matched the specified pattern on the remote machine." |
| SHA-256 hash mismatch | Console displays "WARNING: Hash mismatch detected. File may be corrupt or incomplete." |
| WinRM session drops mid-transfer | Console displays "ERROR: Session lost during transfer. Check network connectivity." |
| Destination directory does not exist | The tool creates the directory automatically before transfer. |

---

## Integration with IT_HealthCheck

IT_RemoteDeploy is designed as a **companion deployment tool** for the IT_HealthCheck diagnostic dashboard.

```
Typical Workflow:

1. [IT_RemoteDeploy - Push]
   Push IT_HealthCheck.ps1 to the remote machine.

2. [IT_HealthCheck - Run]
   Execute the diagnostic dashboard on the remote endpoint.
   Generate capture CSV + HTML reports.

3. [IT_RemoteDeploy - Pull]
   Pull the generated reports back to the local IT workstation.

4. [Copilot Analysis]
   Paste the Copilot Diagnostic Prompt from the HTML report
   into Microsoft Copilot for root cause analysis and
   ServiceNow-ready documentation.
```

```
+------------------+       +------------------+       +------------------+
|  IT_RemoteDeploy |  -->  |  IT_HealthCheck  |  -->  |  IT_RemoteDeploy |
|  (Push Script)   |       |  (Run on Remote) |       |  (Pull Reports)  |
+------------------+       +------------------+       +------------------+
                                                              |
                                                              v
                                                      +------------------+
                                                      | Copilot Analysis |
                                                      | (Root Cause +    |
                                                      |  Ticket Docs)    |
                                                      +------------------+
```

---

## Design Principles

| Principle | Description |
|---|---|
| **Evidence-First** | Every transfer is SHA-256 verified. Trust the hash, not assumptions. |
| **Audit-Ready** | Every action is timestamped and logged. Console output can be screenshot or copied for tickets. |
| **Non-Destructive** | The tool only copies files. No deletions, no modifications, no remote code execution. |
| **Safe by Default** | Connection test must pass before transfers are allowed. No blind pushes or pulls. |
| **UI-Responsive** | Background runspaces keep the GUI alive during long network operations. |
| **General IT Use** | Designed for any IT department, not limited to executive support. |

---

## Pro Tips

- **Always run the Connection Test first.** It catches 90% of issues before you attempt a transfer.
- **Use specific wildcard patterns** in Pull mode to avoid downloading unrelated files.
- **Check the SHA-256 hashes** in the console log if you suspect network corruption or incomplete transfers.
- **Combine with IT_HealthCheck:** Push the script, run the dashboard, pull the CSV/HTML reports -- all without leaving your desk.
- **Keep C:\IT_Scripts as the standard remote directory** for consistency across all managed endpoints.
- **Screenshot the console log** after each operation for ServiceNow work notes or audit documentation.
- **If WinRM fails**, verify the remote machine has PS Remoting enabled and the Windows Firewall allows WinRM traffic (TCP 5985/5986).

---

## File Structure

```
C:\IT_Scripts\
    |
    |-- IT_RemoteDeploy.ps1        # Main deployment tool (this script)
    |-- IT_HealthCheck.ps1         # Companion diagnostic dashboard
    |
C:\Temp\
    |
    |-- <HOSTNAME>_capture_log_<timestamp>.csv    # Pulled capture logs
    |-- <HOSTNAME>_<timestamp>.html               # Pulled HTML dashboards
```

---

## License

This tool is intended for **internal IT department use only**. Not for public distribution or resale.

---

<p align="center">
  <b>IT_RemoteDeploy V1.0</b><br>
  Secure File Transfer for IT Support Teams<br>
  <i>Push. Pull. Verify. Document.</i>
</p>
