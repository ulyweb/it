# Windows 11 Maintenance Dashboard

A dark-themed, animated browser-based maintenance console for Windows 11.

## Files
- `dashboard.html` — The GUI (frontend)
- `launcher.py` — Python backend server that executes commands
- `START_DASHBOARD.bat` — Double-click launcher (easiest way to start)

## Requirements
- Python 3.7+ (free from https://python.org)
  → During install: check ✅ "Add Python to PATH"

## How to Run

### Option A — Double-click (easiest)
Right-click `START_DASHBOARD.bat` → **Run as Administrator**
The dashboard opens automatically in your browser.

### Option B — Command line
Open Command Prompt as **Administrator**, then:
```
cd path\to\this\folder
python launcher.py
```
Then open: http://localhost:9191

## Commands Included

| # | Command | Requires Admin | Requires Reboot |
|---|---------|:-:|:-:|
| 1 | winget upgrade --all | ✅ | ❌ |
| 2 | sfc /scannow | ✅ | ❌ |
| 3 | DISM /RestoreHealth | ✅ | ❌ |
| 4 | chkdsk /r | ✅ | ✅ |
| 5 | ipconfig /flushdns + release + renew | ✅ | ❌ |
| 6 | netsh winsock/ip/tcp reset | ✅ | ✅ |
| 7 | Get-PhysicalDisk | ❌ | ❌ |
| 8 | tasklist + taskkill | ❌ | ❌ |
| 9 | powercfg /energy | ✅ | ❌ |
| 10 | mdsched | ✅ | ✅ |

## Features
- Run commands individually or all in sequence
- Live streaming output in a terminal window
- Process list with kill button
- Reboot warning for commands that need it
- Copy output to clipboard
- Dark animated UI with status indicators

## Notes
- Some commands (sfc, DISM, chkdsk) may take 10-30+ minutes
- powercfg /energy report saves to your Desktop as `energy_report.html`
- The "Run All" button queues all 10 commands in order
- Stop button cancels the current run
