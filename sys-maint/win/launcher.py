"""
Windows 11 Maintenance Dashboard - Backend Launcher
Run: python launcher.py
Then open: http://localhost:9191
"""

import http.server
import socketserver
import subprocess
import json
import os
import sys
import threading
import webbrowser
import time
from urllib.parse import urlparse, parse_qs

PORT = 9191
DASHBOARD_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dashboard.html")

# Registry keys required for winget to work under policy restrictions
WINGET_REG_BASE = r"HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\AppInstaller"
WINGET_REG_KEYS = [
    ("EnableAppInstaller",           "REG_DWORD", "1"),
    ("EnableExperimentalFeatures",   "REG_DWORD", "1"),
    ("EnableHashOverride",           "REG_DWORD", "1"),
    ("EnableMSAppInstallerProtocol", "REG_DWORD", "1"),
]

# ── Browser Fix ─────────────────────────────────────────────────────────────
# Key registry checks we use as a quick "is it already applied?" pre-flight.
# We only check a representative subset — enough to know whether the script
# has been run before. Full apply always writes every key.
BROWSER_CHECK_KEYS = [
    # (hive, path, value_name, expected_data_contains)
    ("HKLM", r"SOFTWARE\Policies\Google\Update",  "AutoUpdateCheckPeriodMinutes", "0x1"),
    ("HKLM", r"SOFTWARE\Policies\Google\Chrome",  "DownloadRestrictions",         "0x0"),
    ("HKLM", r"SOFTWARE\Policies\Microsoft\Edge", "BrowserSignin",                "0x1"),
    ("HKLM", r"SOFTWARE\Policies\Microsoft\Edge", "SyncDisabled",                 "0x0"),
]

def check_browser_registry():
    """Check whether browser policy registry keys look correct."""
    missing = []
    present = []
    for (hive, path, name, expected) in BROWSER_CHECK_KEYS:
        full_path = f"{hive}\\{path}"
        try:
            result = subprocess.run(
                ["reg", "query", full_path, "/v", name],
                capture_output=True, text=True
            )
            if result.returncode == 0 and expected in result.stdout.lower():
                present.append(name)
            else:
                missing.append(name)
        except Exception:
            missing.append(name)
    return {
        "all_present": len(missing) == 0,
        "missing": missing,
        "present": present
    }

# The corrected/hardened PowerShell script (inline, no external .ps1 file needed)
BROWSER_FIX_PS1 = r"""
$ErrorActionPreference = 'Stop'
$log = @()
function L($m) { Write-Output $m; $script:log += $m }

L "=== Browser Policy Fix ==="
L "Started: $(Get-Date)"

# ── Helper: ensure registry path exists ──────────────────────────────────────
function Ensure-Path($path) {
    if (-not (Test-Path $path)) {
        New-Item -Path $path -Force | Out-Null
        Write-Output "  Created: $path"
    }
}

# ── Google Chrome paths ───────────────────────────────────────────────────────
$chromeUpdate = "HKLM:\SOFTWARE\Policies\Google\Update"
$chrome       = "HKLM:\SOFTWARE\Policies\Google\Chrome"

Ensure-Path $chromeUpdate
Ensure-Path $chrome

L "--- Chrome: Auto-Update ---"
New-ItemProperty -Path $chromeUpdate -Name "AutoUpdateCheckPeriodMinutes"              -PropertyType DWord  -Value 1 -Force | Out-Null
New-ItemProperty -Path $chromeUpdate -Name "Install{8A69D345-D564-463C-AFF1-A69D9E530F96}" -PropertyType DWord -Value 1 -Force | Out-Null
New-ItemProperty -Path $chromeUpdate -Name "Update{8A69D345-D564-463C-AFF1-A69D9E530F96}"  -PropertyType DWord -Value 1 -Force | Out-Null
L "  Chrome auto-update enabled."

L "--- Chrome: Downloads & Safe Browsing ---"
Remove-ItemProperty -Path $chrome -Name "DisableSafeBrowsing" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $chrome -Name "BlockThirdPartyCookies" -ErrorAction SilentlyContinue
New-ItemProperty -Path $chrome -Name "DownloadRestrictions" -PropertyType DWord -Value 0 -Force | Out-Null
L "  Downloads unrestricted, Safe Browsing policy removed."

L "--- Chrome: Autofill & Sync ---"
New-ItemProperty -Path $chrome -Name "AutofillAddressEnabled"    -PropertyType DWord -Value 1 -Force | Out-Null
New-ItemProperty -Path $chrome -Name "AutofillCreditCardEnabled" -PropertyType DWord -Value 1 -Force | Out-Null
New-ItemProperty -Path $chrome -Name "ImportAutofillFormData"    -PropertyType DWord -Value 1 -Force | Out-Null
New-ItemProperty -Path $chrome -Name "SyncDisabled"              -PropertyType DWord -Value 0 -Force | Out-Null
L "  Chrome autofill and sync enabled."

# ── Microsoft Edge paths ──────────────────────────────────────────────────────
$edgeHKLM = "HKLM:\SOFTWARE\Policies\Microsoft\Edge"
$edgeHKCU = "HKCU:\Software\Policies\Microsoft\Edge"

Ensure-Path $edgeHKLM
Ensure-Path $edgeHKCU

L "--- Edge: Sign-in & Sync ---"
New-ItemProperty -Path $edgeHKLM -Name "BrowserSignin"        -PropertyType DWord  -Value 1  -Force | Out-Null
New-ItemProperty -Path $edgeHKLM -Name "ForceSync"            -PropertyType DWord  -Value 1  -Force | Out-Null
New-ItemProperty -Path $edgeHKLM -Name "SyncDisabled"         -PropertyType DWord  -Value 0  -Force | Out-Null
New-ItemProperty -Path $edgeHKLM -Name "SyncTypesListDisabled"-PropertyType String -Value "" -Force | Out-Null
L "  Edge sign-in and sync enabled."

L "--- Edge: Password Manager ---"
New-ItemProperty -Path $edgeHKLM -Name "PasswordManagerEnabled" -PropertyType DWord -Value 1 -Force | Out-Null
New-ItemProperty -Path $edgeHKLM -Name "PasswordMonitorAllowed" -PropertyType DWord -Value 1 -Force | Out-Null
L "  Edge password manager enabled."

L "--- Edge: Autofill & Payments ---"
New-ItemProperty -Path $edgeHKCU -Name "AutofillAddressEnabled"     -PropertyType DWord -Value 1 -Force | Out-Null
New-ItemProperty -Path $edgeHKLM -Name "AutofillAddressEnabled"     -PropertyType DWord -Value 1 -Force | Out-Null
New-ItemProperty -Path $edgeHKLM -Name "AutofillPredictionsEnabled" -PropertyType DWord -Value 1 -Force | Out-Null
New-ItemProperty -Path $edgeHKLM -Name "AutofillCreditCardEnabled"  -PropertyType DWord -Value 1 -Force | Out-Null
New-ItemProperty -Path $edgeHKLM -Name "PaymentMethodQueryEnabled"  -PropertyType DWord -Value 1 -Force | Out-Null
L "  Edge autofill and payment info enabled."

L "=== Completed: $(Get-Date) ==="
"""

def check_winget_registry():
    """Check which winget policy registry keys are present/correct."""
    missing = []
    present = []
    for (name, _type, _val) in WINGET_REG_KEYS:
        try:
            result = subprocess.run(
                ["reg", "query", WINGET_REG_BASE, "/v", name],
                capture_output=True, text=True
            )
            if result.returncode == 0 and "0x1" in result.stdout.lower():
                present.append(name)
            else:
                missing.append(name)
        except Exception:
            missing.append(name)
    return {
        "all_present": len(missing) == 0,
        "missing": missing,
        "present": present
    }

# Command definitions
COMMANDS = {
    "browser_fix": {
        "name": "Browser Policy Fix (Chrome + Edge)",
        "requires_reboot": False,
        "requires_admin": True,
        "steps": [
            {
                "cmd": ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                        "-Command", BROWSER_FIX_PS1],
                "label": "Applying Chrome & Edge registry policies..."
            }
        ]
    },
    "winget_unlock": {
        "name": "Unlock winget (Registry Fix)",
        "requires_reboot": False,
        "requires_admin": True,
        "steps": [
            {"cmd": ["REG", "ADD", WINGET_REG_BASE, "/v", "EnableAppInstaller",           "/t", "REG_DWORD", "/d", "1", "/f"], "label": "Enabling AppInstaller policy..."},
            {"cmd": ["REG", "ADD", WINGET_REG_BASE, "/v", "EnableExperimentalFeatures",   "/t", "REG_DWORD", "/d", "1", "/f"], "label": "Enabling Experimental Features..."},
            {"cmd": ["REG", "ADD", WINGET_REG_BASE, "/v", "EnableHashOverride",           "/t", "REG_DWORD", "/d", "1", "/f"], "label": "Enabling Hash Override..."},
            {"cmd": ["REG", "ADD", WINGET_REG_BASE, "/v", "EnableMSAppInstallerProtocol", "/t", "REG_DWORD", "/d", "1", "/f"], "label": "Enabling MS AppInstaller Protocol..."},
        ]
    },
    "winget": {
        "name": "Windows Update (winget)",
        "requires_reboot": False,
        "requires_admin": True,
        "steps": [
            {"cmd": ["winget", "upgrade", "--all", "--accept-source-agreements", "--accept-package-agreements"], "label": "Upgrading all packages..."}
        ]
    },
    "sfc": {
        "name": "System File Checker",
        "requires_reboot": False,
        "requires_admin": True,
        "steps": [
            {"cmd": ["sfc", "/scannow"], "label": "Scanning system files..."}
        ]
    },
    "dism": {
        "name": "DISM Repair",
        "requires_reboot": False,
        "requires_admin": True,
        "steps": [
            {"cmd": ["DISM", "/Online", "/Cleanup-Image", "/RestoreHealth"], "label": "Restoring system image health..."}
        ]
    },
    "chkdsk": {
        "name": "Check Disk",
        "requires_reboot": True,
        "requires_admin": True,
        "steps": [
            {"cmd": ["cmd", "/c", "echo Y | chkdsk C: /r /f"], "label": "Scheduling disk check (requires reboot)..."}
        ]
    },
    "dns": {
        "name": "Flush DNS & Renew IP",
        "requires_reboot": False,
        "requires_admin": True,
        "steps": [
            {"cmd": ["ipconfig", "/flushdns"], "label": "Flushing DNS cache..."},
            {"cmd": ["ipconfig", "/release"], "label": "Releasing IP address..."},
            {"cmd": ["ipconfig", "/renew"], "label": "Renewing IP address..."}
        ]
    },
    "network": {
        "name": "Network Stack Reset",
        "requires_reboot": True,
        "requires_admin": True,
        "steps": [
            {"cmd": ["netsh", "winsock", "reset"], "label": "Resetting Winsock catalog..."},
            {"cmd": ["netsh", "int", "ip", "reset"], "label": "Resetting IP stack..."},
            {"cmd": ["netsh", "int", "tcp", "reset"], "label": "Resetting TCP stack..."}
        ]
    },
    "disk_health": {
        "name": "Disk Health Status",
        "requires_reboot": False,
        "requires_admin": False,
        "steps": [
            {"cmd": ["powershell", "-Command", "Get-PhysicalDisk | Format-Table FriendlyName, MediaType, HealthStatus, OperationalStatus -AutoSize"], "label": "Querying physical disk status..."}
        ]
    },
    "tasklist": {
        "name": "Task Manager",
        "requires_reboot": False,
        "requires_admin": False,
        "steps": [
            {"cmd": ["tasklist", "/FI", "STATUS eq RUNNING", "/FO", "CSV"], "label": "Listing running processes..."}
        ]
    },
    "powercfg": {
        "name": "Power Efficiency Report",
        "requires_reboot": False,
        "requires_admin": True,
        "steps": [
            {"cmd": ["powercfg", "/energy", "/output", os.path.join(os.path.expanduser("~"), "Desktop", "energy_report.html")], "label": "Generating energy efficiency report..."}
        ]
    },
    "mdsched": {
        "name": "Memory Diagnostic",
        "requires_reboot": True,
        "requires_admin": True,
        "steps": [
            {"cmd": ["mdsched"], "label": "Launching Windows Memory Diagnostic..."}
        ]
    }
}

class SSEOutput:
    def __init__(self, wfile):
        self.wfile = wfile

    def send(self, data):
        try:
            msg = f"data: {json.dumps(data)}\n\n"
            self.wfile.write(msg.encode("utf-8"))
            self.wfile.flush()
        except:
            pass

def run_command_stream(cmd_key, sse):
    if cmd_key not in COMMANDS:
        sse.send({"type": "error", "text": f"Unknown command: {cmd_key}"})
        return

    task = COMMANDS[cmd_key]
    sse.send({"type": "start", "name": task["name"]})

    if task.get("requires_admin"):
        sse.send({"type": "info", "text": "⚠ Admin privileges required. If output is empty, re-run launcher as Administrator."})

    # Auto-check winget registry before running winget upgrade
    if cmd_key == "winget":
        sse.send({"type": "step", "text": "Pre-flight: Checking winget policy registry keys..."})
        reg_status = check_winget_registry()
        if reg_status["all_present"]:
            sse.send({"type": "success", "text": "✓ All winget registry keys are present — proceeding."})
        else:
            sse.send({"type": "warning", "text": f"⚠ Missing registry keys: {', '.join(reg_status['missing'])}"})
            sse.send({"type": "warning", "text": "  winget may be blocked by policy. Applying registry fix automatically..."})
            # Apply the fix inline before running winget
            for (name, rtype, val) in WINGET_REG_KEYS:
                fix_cmd = ["REG", "ADD", WINGET_REG_BASE, "/v", name, "/t", rtype, "/d", val, "/f"]
                sse.send({"type": "cmd", "text": " ".join(fix_cmd)})
                try:
                    r = subprocess.run(fix_cmd, capture_output=True, text=True)
                    if r.returncode == 0:
                        sse.send({"type": "success", "text": f"  ✓ Set {name}"})
                    else:
                        sse.send({"type": "error", "text": f"  ✗ Failed to set {name}: {r.stderr.strip()}"})
                except Exception as ex:
                    sse.send({"type": "error", "text": f"  ✗ Error setting {name}: {str(ex)}"})
            sse.send({"type": "info", "text": "Registry fix applied — continuing with winget upgrade..."})

    for step in task["steps"]:
        sse.send({"type": "step", "text": step["label"]})
        sse.send({"type": "cmd", "text": " ".join(step["cmd"])})
        try:
            proc = subprocess.Popen(
                step["cmd"],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
            )
            for line in proc.stdout:
                line = line.rstrip()
                if line:
                    sse.send({"type": "output", "text": line})
            proc.wait()
            rc = proc.returncode
            if rc == 0:
                sse.send({"type": "success", "text": f"✓ Completed with exit code {rc}"})
            else:
                sse.send({"type": "warning", "text": f"⚠ Exited with code {rc} (some codes are normal, e.g. chkdsk=3)"})
        except FileNotFoundError:
            sse.send({"type": "error", "text": f"Command not found: {step['cmd'][0]}. Make sure it's available in PATH."})
        except Exception as e:
            sse.send({"type": "error", "text": f"Error: {str(e)}"})

    if task.get("requires_reboot"):
        sse.send({"type": "reboot", "text": "⟳ A system reboot is recommended to complete this operation."})

    sse.send({"type": "done", "name": task["name"]})


def handle_taskkill(pid, sse):
    sse.send({"type": "start", "name": f"Kill Process PID {pid}"})
    try:
        proc = subprocess.run(
            ["taskkill", "/PID", str(pid), "/F"],
            capture_output=True, text=True
        )
        out = proc.stdout + proc.stderr
        for line in out.strip().splitlines():
            sse.send({"type": "output", "text": line})
        if proc.returncode == 0:
            sse.send({"type": "success", "text": f"✓ Process {pid} terminated."})
        else:
            sse.send({"type": "error", "text": f"Failed to kill PID {pid}"})
    except Exception as e:
        sse.send({"type": "error", "text": str(e)})
    sse.send({"type": "done", "name": "taskkill"})


class Handler(http.server.BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass  # Suppress default logging

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        qs = parse_qs(parsed.query)

        if path == "/" or path == "/index.html":
            self.serve_file(DASHBOARD_FILE, "text/html")

        elif path == "/api/commands":
            data = {k: {"name": v["name"], "requires_reboot": v["requires_reboot"], "requires_admin": v["requires_admin"]} for k, v in COMMANDS.items()}
            self.send_json(data)

        elif path == "/api/run":
            cmd_key = qs.get("cmd", [""])[0]
            pid = qs.get("pid", [""])[0]

            self.send_response(200)
            self.send_header("Content-Type", "text/event-stream")
            self.send_header("Cache-Control", "no-cache")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()

            sse = SSEOutput(self.wfile)

            if pid:
                handle_taskkill(pid, sse)
            elif cmd_key:
                run_command_stream(cmd_key, sse)
            else:
                sse.send({"type": "error", "text": "No command specified."})

        elif path == "/api/shutdown":
            self.send_json({"status": "shutting_down"})
            # Graceful shutdown in a background thread so the response gets sent first
            def _stop():
                time.sleep(0.3)
                print("\n  Shutdown requested from dashboard. Stopping server...")
                self.server.shutdown()
            threading.Thread(target=_stop, daemon=True).start()

        elif path == "/api/check_browser":
            self.send_json(check_browser_registry())

        elif path == "/api/check_winget":
            self.send_json(check_winget_registry())

        elif path == "/api/status":
            self.send_json({"status": "ok", "port": PORT})

        else:
            self.send_response(404)
            self.end_headers()

    def serve_file(self, filepath, content_type):
        try:
            with open(filepath, "rb") as f:
                content = f.read()
            self.send_response(200)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", len(content))
            self.end_headers()
            self.wfile.write(content)
        except FileNotFoundError:
            self.send_response(404)
            self.end_headers()
            self.wfile.write(b"File not found")

    def send_json(self, data):
        body = json.dumps(data).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Content-Length", len(body))
        self.end_headers()
        self.wfile.write(body)


def main():
    print("=" * 55)
    print("  Windows 11 Maintenance Dashboard")
    print("=" * 55)
    print(f"  Server: http://localhost:{PORT}")
    print("  TIP: Run as Administrator for full functionality")
    print("  Press Ctrl+C to stop")
    print("=" * 55)

    with socketserver.TCPServer(("", PORT), Handler) as httpd:
        httpd.allow_reuse_address = True
        threading.Thread(target=lambda: (time.sleep(1.2), webbrowser.open(f"http://localhost:{PORT}")), daemon=True).start()
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\n  Server stopped.")


if __name__ == "__main__":
    main()
