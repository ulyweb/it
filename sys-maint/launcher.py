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

# Command definitions
COMMANDS = {
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
