### Usage

Running it
````
powershell.exe -ExecutionPolicy Bypass -File .\WiFi-Connectivity-Monitor-v3.ps1
````

---

WinForms dialog (dark/cyan, matching the theme) asks for capture point and depth; falls back to a console menu on Server Core. No admin required. Non-interactive for a fleet push:

````
.\WiFi-Connectivity-Monitor-v3.ps1 -Location "Wired Ethernet Test" -Samples 180 -NoPrompt
````

Output: C:\Temp\<timestamp>_WiFiConnectivityMonitor_<location>\ — HTML dashboard, PingSamples.csv, CaptureLog.log, NetshRaw.txt.
