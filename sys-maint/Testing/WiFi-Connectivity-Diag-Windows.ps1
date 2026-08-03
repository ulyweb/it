<#
.SYNOPSIS
  WiFi Connectivity Diagnostic Capture - Windows
.DESCRIPTION
  Read-only diagnostic tool for Executive Support. Captures Wi-Fi state, IP configuration,
  DNS/gateway reachability, sustained ping results, route/traceroute, profile inventory,
  and generates a dark/gold telemetry HTML dashboard plus CSV/log outputs.
.NOTES
  Run in PowerShell as the signed-in user. Admin is not required for normal capture.
#>

param(
    [string]$PrimaryTarget = "8.8.8.8",
    [string]$DnsTarget = "google.com",
    [int]$PingCount = 60,
    [int]$SampleIntervalSeconds = 1
)

$ErrorActionPreference = "Continue"
$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$outDir = "C:\Temp\${ts}_WiFiConnectivityDiag"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$logPath = Join-Path $outDir "WiFiConnectivityDiag.log"
$htmlPath = Join-Path $outDir "WiFiConnectivityDiag_Report.html"
$pingCsv = Join-Path $outDir "PingSamples.csv"
$adapterCsv = Join-Path $outDir "AdapterSummary.csv"

function Write-Log {
    param([string]$Message)
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Add-Content -Path $logPath -Value $line
}

function Run-Cmd {
    param([string]$Command)
    Write-Log "RUN: $Command"
    try {
        $output = cmd.exe /c $Command 2>&1
        return ($output -join "`r`n")
    } catch {
        return "ERROR running $Command : $($_.Exception.Message)"
    }
}

function HtmlEncode {
    param([string]$Text)
    if ($null -eq $Text) { return "" }
    return [System.Net.WebUtility]::HtmlEncode($Text)
}

Write-Log "Starting Windows WiFi Connectivity Diagnostic Capture"

$computer = $env:COMPUTERNAME
$user = $env:USERNAME
$os = Get-CimInstance Win32_OperatingSystem
$bios = Get-CimInstance Win32_BIOS
$model = Get-CimInstance Win32_ComputerSystem

$wifiInterfaceRaw = Run-Cmd "netsh wlan show interfaces"
$wifiDriversRaw = Run-Cmd "netsh wlan show drivers"
$wifiProfilesRaw = Run-Cmd "netsh wlan show profiles"
$ipconfigRaw = Run-Cmd "ipconfig /all"
$routeRaw = Run-Cmd "route print"
$dnsRaw = Run-Cmd "nslookup $DnsTarget"
$traceRaw = Run-Cmd "tracert -d -h 12 $PrimaryTarget"

$wifiAdapters = Get-NetAdapter -Physical -ErrorAction SilentlyContinue | Where-Object { $_.NdisPhysicalMedium -match "Native802_11|WirelessLan|Unspecified" -or $_.InterfaceDescription -match "Wi-Fi|Wireless|802.11" }
$netIpConfig = Get-NetIPConfiguration -ErrorAction SilentlyContinue
$dnsServers = Get-DnsClientServerAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue

$gateway = ($netIpConfig | Where-Object { $_.IPv4DefaultGateway -and $_.NetAdapter.Status -eq "Up" } | Select-Object -First 1).IPv4DefaultGateway.NextHop
if (-not $gateway) { $gateway = "Not detected" }

$pingRows = New-Object System.Collections.Generic.List[object]
Write-Log "Starting ping sample test. Target=$PrimaryTarget Count=$PingCount Interval=$SampleIntervalSeconds"
for ($i = 1; $i -le $PingCount; $i++) {
    $start = Get-Date
    $result = Test-Connection -ComputerName $PrimaryTarget -Count 1 -ErrorAction SilentlyContinue
    if ($result) {
        $latency = if ($result.ResponseTime) { [int]$result.ResponseTime } elseif ($result.Latency) { [int]$result.Latency } else { $null }
        $status = "Success"
    } else {
        $latency = $null
        $status = "Failed"
    }
    $pingRows.Add([pscustomobject]@{
        Sample = $i
        Timestamp = $start.ToString("yyyy-MM-dd HH:mm:ss")
        Target = $PrimaryTarget
        Status = $status
        LatencyMs = $latency
    })
    Start-Sleep -Seconds $SampleIntervalSeconds
}
$pingRows | Export-Csv -NoTypeInformation -Path $pingCsv

$successCount = ($pingRows | Where-Object Status -eq "Success").Count
$failCount = ($pingRows | Where-Object Status -eq "Failed").Count
$lossPct = if ($PingCount -gt 0) { [math]::Round(($failCount / $PingCount) * 100, 2) } else { 0 }
$avgLatency = ($pingRows | Where-Object { $_.LatencyMs -ne $null } | Measure-Object -Property LatencyMs -Average).Average
$maxLatency = ($pingRows | Where-Object { $_.LatencyMs -ne $null } | Measure-Object -Property LatencyMs -Maximum).Maximum
$avgLatency = if ($avgLatency) { [math]::Round($avgLatency, 2) } else { "N/A" }
$maxLatency = if ($maxLatency) { [math]::Round($maxLatency, 2) } else { "N/A" }

$gatewayPing = if ($gateway -and $gateway -ne "Not detected") { Test-Connection -ComputerName $gateway -Count 4 -ErrorAction SilentlyContinue } else { $null }
$gatewayStatus = if ($gatewayPing) { "Reachable" } else { "Not reachable / not detected" }

$adapterRows = foreach ($a in (Get-NetAdapter -ErrorAction SilentlyContinue)) {
    [pscustomobject]@{
        Name = $a.Name
        InterfaceDescription = $a.InterfaceDescription
        Status = $a.Status
        LinkSpeed = $a.LinkSpeed
        MacAddress = $a.MacAddress
    }
}
$adapterRows | Export-Csv -NoTypeInformation -Path $adapterCsv

$health = "GREEN"
$healthText = "No packet loss detected during sampling"
if ($lossPct -gt 0 -and $lossPct -lt 10) { $health = "AMBER"; $healthText = "Intermittent packet loss detected" }
if ($lossPct -ge 10) { $health = "RED"; $healthText = "Significant packet loss detected" }

$recommendations = New-Object System.Collections.Generic.List[string]
if ($lossPct -gt 0) { $recommendations.Add("Packet loss was observed. Compare this result against a wired Ethernet test to isolate Wi-Fi versus ISP/router path.") }
if ($gatewayStatus -notlike "Reachable*") { $recommendations.Add("Default gateway did not respond. Validate local router/AP, DHCP, and network path.") }
$recommendations.Add("Run this same capture near the access point and again at the affected desk location to compare signal and packet loss.")
$recommendations.Add("If wired Ethernet is stable while Wi-Fi drops, focus on RF interference, AP placement, band/channel congestion, and office layout.")
$recommendations.Add("If wired and wireless both fail, focus on Comcast Business service, modem/router health, firewall, DHCP, and upstream packet loss.")

$style = @"
<style>
body{margin:0;background:#070707;color:#f5f2e9;font-family:'Segoe UI',Arial,sans-serif;}
.bg{position:fixed;inset:0;background:radial-gradient(circle at 20% 10%,rgba(212,175,55,.18),transparent 30%),radial-gradient(circle at 80% 30%,rgba(212,175,55,.10),transparent 25%),linear-gradient(135deg,#080808,#131313 45%,#050505);z-index:-1;}
.wrap{max-width:1200px;margin:0 auto;padding:34px;}
.hero{border:1px solid rgba(212,175,55,.35);background:rgba(18,18,18,.72);backdrop-filter:blur(18px);border-radius:24px;padding:28px;box-shadow:0 0 40px rgba(212,175,55,.12);}
h1{margin:0;color:#ffd76a;font-size:32px;letter-spacing:.3px;}.sub{color:#cfc6aa;margin-top:8px;}
.grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:16px;margin:22px 0;}.card{border:1px solid rgba(212,175,55,.25);background:rgba(255,255,255,.05);border-radius:20px;padding:18px;box-shadow:0 14px 40px rgba(0,0,0,.25);}.k{color:#bca75a;font-size:12px;text-transform:uppercase;letter-spacing:1px}.v{font-size:24px;font-weight:700;margin-top:8px}.pill{display:inline-block;padding:6px 12px;border-radius:999px;font-weight:700}.GREEN{background:#073b24;color:#66f0a5}.AMBER{background:#4a3500;color:#ffd36b}.RED{background:#4b0b0b;color:#ff7b7b}pre{white-space:pre-wrap;background:#0d0d0d;border:1px solid rgba(212,175,55,.2);border-radius:16px;padding:16px;overflow:auto;color:#eee}.section{margin-top:24px}.section h2{color:#ffd76a;border-bottom:1px solid rgba(212,175,55,.25);padding-bottom:8px}.small{color:#bfb7a1;font-size:12px}li{margin:8px 0}.two{display:grid;grid-template-columns:1fr 1fr;gap:18px}@media(max-width:900px){.grid,.two{grid-template-columns:1fr}}
</style>
"@

$recHtml = ($recommendations | ForEach-Object { "<li>$(HtmlEncode $_)</li>" }) -join "`n"
$adapterHtml = ($adapterRows | ConvertTo-Html -Fragment | Out-String)
$pingHtml = ($pingRows | Select-Object -Last 20 | ConvertTo-Html -Fragment | Out-String)

$html = @"
<!doctype html><html><head><meta charset='utf-8'><title>WiFi Connectivity Diagnostic Report - Windows</title>$style</head><body><div class='bg'></div><div class='wrap'>
<div class='hero'><h1>Windows WiFi Connectivity Diagnostic Report</h1><div class='sub'>Executive Support telemetry capture for intermittent Wi-Fi, ISP, gateway, and endpoint network validation.</div><p><span class='pill $health'>$health</span> $healthText</p></div>
<div class='grid'>
<div class='card'><div class='k'>Computer</div><div class='v'>$(HtmlEncode $computer)</div></div>
<div class='card'><div class='k'>User</div><div class='v'>$(HtmlEncode $user)</div></div>
<div class='card'><div class='k'>Packet Loss</div><div class='v'>$lossPct%</div></div>
<div class='card'><div class='k'>Avg / Max Latency</div><div class='v'>$avgLatency / $maxLatency ms</div></div>
</div>
<div class='two'>
<div class='card'><div class='k'>OS</div><div>$(HtmlEncode $($os.Caption)) $($os.Version)</div></div>
<div class='card'><div class='k'>Hardware</div><div>$(HtmlEncode $($model.Manufacturer)) $(HtmlEncode $($model.Model))<br>Serial: $(HtmlEncode $($bios.SerialNumber))</div></div>
</div>
<div class='section card'><h2>Recommended Next Actions</h2><ul>$recHtml</ul></div>
<div class='section card'><h2>Ping Sample Summary</h2><p>Target: $(HtmlEncode $PrimaryTarget) | Count: $PingCount | Success: $successCount | Failed: $failCount | Gateway: $(HtmlEncode $gateway) - $(HtmlEncode $gatewayStatus)</p>$pingHtml</div>
<div class='section card'><h2>Adapter Summary</h2>$adapterHtml</div>
<div class='section card'><h2>Wi-Fi Interface</h2><pre>$(HtmlEncode $wifiInterfaceRaw)</pre></div>
<div class='section card'><h2>Wi-Fi Driver</h2><pre>$(HtmlEncode $wifiDriversRaw)</pre></div>
<div class='section card'><h2>Wi-Fi Profiles</h2><pre>$(HtmlEncode $wifiProfilesRaw)</pre></div>
<div class='section card'><h2>IP Configuration</h2><pre>$(HtmlEncode $ipconfigRaw)</pre></div>
<div class='section card'><h2>DNS Test</h2><pre>$(HtmlEncode $dnsRaw)</pre></div>
<div class='section card'><h2>Traceroute</h2><pre>$(HtmlEncode $traceRaw)</pre></div>
<div class='section card'><h2>Route Table</h2><pre>$(HtmlEncode $routeRaw)</pre></div>
<p class='small'>Generated: $(Get-Date) | Output folder: $(HtmlEncode $outDir)</p>
</div></body></html>
"@

$html | Set-Content -Path $htmlPath -Encoding UTF8
Write-Log "Report generated: $htmlPath"
Start-Process $htmlPath
Write-Host "Report generated: $htmlPath"
