<#
================================================================================
 WIFI CONNECTIVITY MONITOR v3.1  --  Windows
 
 Companion to WiFi-Connectivity-Monitor-v3.1 for macOS. Both builds share the
 identical dashboard engine, so a Windows capture and a Mac capture render the
 same panels, the same root cause matrix and the same comparison ledger.

 Interface pattern modeled on worldmonitor.app / PG&E POWER MONITOR v4.0:
   sticky topbar, LIVE pulse, version + status pills, command palette (Ctrl K),
   keyboard shortcuts, dark/light theme, panel chips, collapsible panels,
   KPI strip, zero-dependency canvas charts, toasts, JSON export, print/PDF,
   and a localStorage capture ledger for multi-position comparison.

 NEW IN v3.1 - PLAIN ENGLISH BRIEFING
   A first-class panel written for non-technical readers. Every measurement is
   explained with an everyday analogy and interpreted against the actual
   captured value, plus a jargon decoder, a "what happens next" list and a
   ready-to-send status message with a one-click copy button. A third Copilot
   prompt generates an end-user summary on demand.
   Shortcut: X jumps to it. M copies the status message.

 READ-ONLY capture. No network settings are changed. Admin is NOT required.

 Output: C:\Temp\<timestamp>_WiFiConnectivityMonitor_<location>\
           WiFi_Connectivity_Monitor.html   dashboard
           PingSamples.csv                  full sample set
           CaptureLog.log                   audit log
           NetshRaw.txt                     raw netsh output

 PLATFORM NOTE
   Windows reports wireless signal as a link-quality percentage and does not
   expose a noise floor, so SNR cannot be measured. RSSI is derived from the
   quality percentage using the standard conversion (dBm = pct/2 - 100). The
   dashboard suppresses every SNR-derived gauge, KPI and score on Windows
   rather than fabricating a value.

 USAGE
   powershell.exe -ExecutionPolicy Bypass -File .\WiFi-Connectivity-Monitor-v3.1-Windows.ps1
   powershell.exe -ExecutionPolicy Bypass -File .\WiFi-Connectivity-Monitor-v3.1-Windows.ps1 -Samples 180 -Location "Wired Ethernet Test" -NoPrompt
================================================================================
#>

[CmdletBinding()]
param(
    [string]$PingTarget = "8.8.8.8",
    [string]$DnsTarget  = "google.com",
    [string]$HttpTarget = "https://www.msftconnecttest.com/connecttest.txt",
    [int]$Samples = 0,
    [string]$Location = "",
    [switch]$NoPrompt
)

$ErrorActionPreference = "Continue"
$ProgressPreference    = "Continue"
$TOOL_VERSION          = "3.1"

# ---------------------------------------------------------------- helpers ---

function Invoke-Native {
    param([string]$Command)
    try   { return (cmd.exe /c $Command 2>&1 | Out-String) }
    catch { return "Command did not complete: $($_.Exception.Message)" }
}

function Get-NetshValue {
    param([string]$Block, [string]$Key)
    foreach ($line in ($Block -split "`r?`n")) {
        $idx = $line.IndexOf(':')
        if ($idx -lt 1) { continue }
        $k = $line.Substring(0, $idx).Trim()
        if ($k -ieq $Key) { return $line.Substring($idx + 1).Trim() }
    }
    return ""
}

# Delegate JSON string escaping to ConvertTo-Json so quotes, backslashes and
# control characters in raw command output can never corrupt the payload.
function ConvertTo-JsString {
    param([string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return '""' }
    return ($Value | ConvertTo-Json -Compress)
}

function Write-Log {
    param([string]$Message)
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Add-Content -Path $script:LogPath -Value $line -Encoding UTF8
}

function Round2 { param([double]$v) return [math]::Round($v, 2) }

# ------------------------------------------------------- capture options ---

$locationChoices = @(
    "Desk / User Location",
    "Near Router or Access Point",
    "Wired Ethernet Test",
    "Conference Room",
    "Reception or Lobby",
    "Other Location"
)
$depthChoices = @(
    "Quick - 20 samples (about 20 sec)",
    "Standard - 60 samples (about 1 min)",
    "Extended - 180 samples (about 3 min)",
    "Deep - 300 samples (about 5 min)"
)

$hasForms = $false
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    $hasForms = $true
} catch { $hasForms = $false }

function Show-CaptureDialog {
    $form                 = New-Object System.Windows.Forms.Form
    $form.Text            = "WiFi Connectivity Monitor v$TOOL_VERSION"
    $form.Size            = New-Object System.Drawing.Size(520, 300)
    $form.StartPosition   = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox     = $false
    $form.MinimizeBox     = $false
    $form.BackColor       = [System.Drawing.Color]::FromArgb(11, 17, 27)
    $form.ForeColor       = [System.Drawing.Color]::FromArgb(215, 230, 242)

    $title           = New-Object System.Windows.Forms.Label
    $title.Text      = "WIFI CONNECTIVITY MONITOR"
    $title.Font      = New-Object System.Drawing.Font("Consolas", 13, [System.Drawing.FontStyle]::Bold)
    $title.ForeColor = [System.Drawing.Color]::FromArgb(0, 229, 255)
    $title.Location  = New-Object System.Drawing.Point(20, 16)
    $title.Size      = New-Object System.Drawing.Size(470, 26)
    $form.Controls.Add($title)

    $sub           = New-Object System.Windows.Forms.Label
    $sub.Text      = "Read-only capture. Results are stored and compared in the dashboard."
    $sub.ForeColor = [System.Drawing.Color]::FromArgb(125, 147, 168)
    $sub.Location  = New-Object System.Drawing.Point(20, 42)
    $sub.Size      = New-Object System.Drawing.Size(470, 20)
    $form.Controls.Add($sub)

    $lblLoc          = New-Object System.Windows.Forms.Label
    $lblLoc.Text     = "CAPTURE POINT"
    $lblLoc.Font     = New-Object System.Drawing.Font("Consolas", 8)
    $lblLoc.ForeColor= [System.Drawing.Color]::FromArgb(78, 98, 117)
    $lblLoc.Location = New-Object System.Drawing.Point(20, 78)
    $lblLoc.Size     = New-Object System.Drawing.Size(470, 16)
    $form.Controls.Add($lblLoc)

    $cbLoc               = New-Object System.Windows.Forms.ComboBox
    $cbLoc.DropDownStyle = "DropDownList"
    $cbLoc.Location      = New-Object System.Drawing.Point(20, 96)
    $cbLoc.Size          = New-Object System.Drawing.Size(460, 24)
    foreach ($c in $locationChoices) { [void]$cbLoc.Items.Add($c) }
    $cbLoc.SelectedIndex = 0
    $form.Controls.Add($cbLoc)

    $lblDep          = New-Object System.Windows.Forms.Label
    $lblDep.Text     = "TEST LENGTH  (longer captures catch intermittent drops far more reliably)"
    $lblDep.Font     = New-Object System.Drawing.Font("Consolas", 8)
    $lblDep.ForeColor= [System.Drawing.Color]::FromArgb(78, 98, 117)
    $lblDep.Location = New-Object System.Drawing.Point(20, 134)
    $lblDep.Size     = New-Object System.Drawing.Size(470, 16)
    $form.Controls.Add($lblDep)

    $cbDep               = New-Object System.Windows.Forms.ComboBox
    $cbDep.DropDownStyle = "DropDownList"
    $cbDep.Location      = New-Object System.Drawing.Point(20, 152)
    $cbDep.Size          = New-Object System.Drawing.Size(460, 24)
    foreach ($c in $depthChoices) { [void]$cbDep.Items.Add($c) }
    $cbDep.SelectedIndex = 1
    $form.Controls.Add($cbDep)

    $ok               = New-Object System.Windows.Forms.Button
    $ok.Text          = "START CAPTURE"
    $ok.Location      = New-Object System.Drawing.Point(250, 205)
    $ok.Size          = New-Object System.Drawing.Size(140, 32)
    $ok.BackColor     = [System.Drawing.Color]::FromArgb(0, 229, 255)
    $ok.ForeColor     = [System.Drawing.Color]::FromArgb(5, 8, 13)
    $ok.FlatStyle     = "Flat"
    $ok.DialogResult  = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($ok)
    $form.AcceptButton = $ok

    $cancel              = New-Object System.Windows.Forms.Button
    $cancel.Text         = "Cancel"
    $cancel.Location     = New-Object System.Drawing.Point(400, 205)
    $cancel.Size         = New-Object System.Drawing.Size(80, 32)
    $cancel.FlatStyle    = "Flat"
    $cancel.ForeColor    = [System.Drawing.Color]::FromArgb(125, 147, 168)
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancel)
    $form.CancelButton = $cancel

    $result = $form.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    return @{ Location = $cbLoc.SelectedItem; Depth = $cbDep.SelectedItem }
}

$depthChoice = "Standard - 60 samples (about 1 min)"

if (-not $NoPrompt -and ([string]::IsNullOrWhiteSpace($Location) -or $Samples -le 0)) {
    if ($hasForms) {
        $picked = Show-CaptureDialog
        if ($null -eq $picked) { Write-Host "Capture cancelled."; return }
        if ([string]::IsNullOrWhiteSpace($Location)) { $Location = $picked.Location }
        $depthChoice = $picked.Depth
    } else {
        Write-Host ""
        Write-Host "  WIFI CONNECTIVITY MONITOR v$TOOL_VERSION" -ForegroundColor Cyan
        Write-Host ""
        for ($i = 0; $i -lt $locationChoices.Count; $i++) { Write-Host ("   {0}. {1}" -f ($i + 1), $locationChoices[$i]) }
        $sel = Read-Host "`n  Capture point [1]"
        if ([string]::IsNullOrWhiteSpace($sel)) { $sel = "1" }
        $li = [int]$sel - 1
        if ($li -lt 0 -or $li -ge $locationChoices.Count) { $li = 0 }
        if ([string]::IsNullOrWhiteSpace($Location)) { $Location = $locationChoices[$li] }
        Write-Host ""
        for ($i = 0; $i -lt $depthChoices.Count; $i++) { Write-Host ("   {0}. {1}" -f ($i + 1), $depthChoices[$i]) }
        $sel2 = Read-Host "`n  Test length [2]"
        if ([string]::IsNullOrWhiteSpace($sel2)) { $sel2 = "2" }
        $di = [int]$sel2 - 1
        if ($di -lt 0 -or $di -ge $depthChoices.Count) { $di = 1 }
        $depthChoice = $depthChoices[$di]
    }
}

if ([string]::IsNullOrWhiteSpace($Location)) { $Location = "Desk / User Location" }
if ($Samples -le 0) {
    $Samples = 60
    if ($depthChoice.StartsWith("Quick"))    { $Samples = 20 }
    if ($depthChoice.StartsWith("Extended")) { $Samples = 180 }
    if ($depthChoice.StartsWith("Deep"))     { $Samples = 300 }
}

# --------------------------------------------------------- output folders ---

$stamp     = Get-Date -Format "yyyyMMdd_HHmmss"
$safeLabel = ($Location -replace '[^A-Za-z0-9]', '_')
$outDir    = "C:\Temp\${stamp}_WiFiConnectivityMonitor_${safeLabel}"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$script:LogPath = Join-Path $outDir "CaptureLog.log"
$htmlPath       = Join-Path $outDir "WiFi_Connectivity_Monitor.html"
$csvPath        = Join-Path $outDir "PingSamples.csv"
$netshPath      = Join-Path $outDir "NetshRaw.txt"

Write-Log "WiFi Connectivity Monitor v$TOOL_VERSION started"
Write-Log "Location: $Location  Samples: $Samples  Target: $PingTarget"

# ---------------------------------------------------------- device details ---

$hostName    = $env:COMPUTERNAME
$consoleUser = $env:USERNAME
$os          = Get-CimInstance Win32_OperatingSystem  -ErrorAction SilentlyContinue
$cs          = Get-CimInstance Win32_ComputerSystem   -ErrorAction SilentlyContinue
$bios        = Get-CimInstance Win32_BIOS             -ErrorAction SilentlyContinue
$cpu         = Get-CimInstance Win32_Processor        -ErrorAction SilentlyContinue | Select-Object -First 1

$modelName = if ($cs)   { "$($cs.Manufacturer) $($cs.Model)".Trim() } else { "Unknown" }
$modelID   = if ($cs)   { "$($cs.SystemFamily)".Trim() }             else { "" }
if ([string]::IsNullOrWhiteSpace($modelID) -and $cs) { $modelID = "$($cs.SystemSKUNumber)".Trim() }
$chipName  = if ($cpu)  { "$($cpu.Name)".Trim() }                    else { "Unknown" }
$serialNum = if ($bios) { "$($bios.SerialNumber)".Trim() }           else { "Unknown" }
$osName    = if ($os)   { "$($os.Caption)".Trim() }                  else { "Windows" }
$osVer     = if ($os)   { "$($os.Version)" }                         else { "" }
$osBuild   = if ($os)   { "$($os.BuildNumber)" }                     else { "" }

# --------------------------------------------------------------- interfaces ---

$wlanIfaceRaw   = Invoke-Native "netsh wlan show interfaces"
$wlanDriverRaw  = Invoke-Native "netsh wlan show drivers"
$wlanProfileRaw = Invoke-Native "netsh wlan show profiles"
$wlanNetRaw     = Invoke-Native "netsh wlan show networks mode=bssid"
$ipconfigRaw    = Invoke-Native "ipconfig /all"
Set-Content -Path $netshPath -Value ($wlanIfaceRaw + "`r`n" + $wlanDriverRaw + "`r`n" + $wlanNetRaw) -Encoding UTF8

$activeCfg = Get-NetIPConfiguration -ErrorAction SilentlyContinue |
             Where-Object { $_.IPv4DefaultGateway -and $_.NetAdapter.Status -eq 'Up' } |
             Select-Object -First 1

$routerIP = "Not detected"
$activeIP = "Not detected"
$ifaceNm  = "Unknown"
$linkType = "Wired or Other"

if ($activeCfg) {
    $routerIP = $activeCfg.IPv4DefaultGateway.NextHop
    if ($activeCfg.IPv4Address) { $activeIP = ($activeCfg.IPv4Address | Select-Object -First 1).IPAddress }
    $ifaceNm  = $activeCfg.InterfaceAlias
    $desc     = "$($activeCfg.InterfaceDescription)"
    if ($ifaceNm -match 'Wi-?Fi|Wireless' -or $desc -match 'Wi-?Fi|Wireless|802\.11') { $linkType = "Wi-Fi" }
}

$wifiAdapter = Get-NetAdapter -ErrorAction SilentlyContinue |
               Where-Object { $_.InterfaceDescription -match 'Wi-?Fi|Wireless|802\.11' } |
               Select-Object -First 1
$wifiDev = if ($wifiAdapter) { $wifiAdapter.Name } else { "Wi-Fi" }

$dnsServers = (Get-DnsClientServerAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
               Where-Object { $_.ServerAddresses.Count -gt 0 -and $_.InterfaceAlias -eq $ifaceNm } |
               Select-Object -ExpandProperty ServerAddresses -ErrorAction SilentlyContinue) -join ', '
if ([string]::IsNullOrWhiteSpace($dnsServers)) {
    $dnsServers = (Get-DnsClientServerAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                   Select-Object -ExpandProperty ServerAddresses -ErrorAction SilentlyContinue |
                   Select-Object -Unique) -join ', '
}
if ([string]::IsNullOrWhiteSpace($dnsServers)) { $dnsServers = "Not detected" }

# ------------------------------------------------------ wireless telemetry ---

$ssidName  = Get-NetshValue $wlanIfaceRaw "SSID"
$bssidVal  = Get-NetshValue $wlanIfaceRaw "BSSID"
$phyMode   = Get-NetshValue $wlanIfaceRaw "Radio type"
$wChannel  = Get-NetshValue $wlanIfaceRaw "Channel"
$wSecurity = Get-NetshValue $wlanIfaceRaw "Authentication"
$txRateRaw = Get-NetshValue $wlanIfaceRaw "Transmit rate (Mbps)"
$signalRaw = Get-NetshValue $wlanIfaceRaw "Signal"

if ([string]::IsNullOrWhiteSpace($ssidName))  { $ssidName  = "Not available" }
if ([string]::IsNullOrWhiteSpace($bssidVal))  { $bssidVal  = "N/A" }
if ([string]::IsNullOrWhiteSpace($phyMode))   { $phyMode   = "N/A" }
if ([string]::IsNullOrWhiteSpace($wChannel))  { $wChannel  = "N/A" }
if ([string]::IsNullOrWhiteSpace($wSecurity)) { $wSecurity = "N/A" }
$txRate = if ([string]::IsNullOrWhiteSpace($txRateRaw)) { "N/A" } else { "$txRateRaw Mbps" }

# Windows exposes link quality as a percentage, not dBm, and never exposes a
# noise floor. Derive RSSI with the standard conversion and flag SNR as absent.
$sigPct    = $null
$rssiVal   = 0
$haveSig   = $false
if ($signalRaw -match '(\d+)') {
    $sigPct  = [int]$Matches[1]
    $rssiVal = [int](($sigPct / 2) - 100)
    $haveSig = $true
}

# neighbouring network density - the evidence for multi-tenant interference
$nbChannels = @()
foreach ($line in ($wlanNetRaw -split "`r?`n")) {
    if ($line -match '^\s*Channel\s*:\s*(\d+)') { $nbChannels += [int]$Matches[1] }
}
$neighborTotal = $nbChannels.Count
$nbJsParts = @()
foreach ($grp in ($nbChannels | Group-Object | Sort-Object Count -Descending | Select-Object -First 20)) {
    $nbJsParts += ("{{ch:{0},n:{1}}}" -f [int]$grp.Name, [int]$grp.Count)
}
$nbJs = "[" + ($nbJsParts -join ",") + "]"

Write-Log "Interface=$ifaceNm Link=$linkType SSID=$ssidName Channel=$wChannel Signal=$signalRaw Neighbours=$neighborTotal"

# ---------------------------------------------------------- stability test ---

$pinger  = New-Object System.Net.NetworkInformation.Ping
$csvRows = New-Object System.Collections.Generic.List[string]
$jsRows  = New-Object System.Collections.Generic.List[string]
$csvRows.Add("Sample,Timestamp,Target,Status,LatencyMs")

$successCount  = 0
$failCount     = 0
$latSum        = 0.0
$latMax        = 0.0
$latMin        = 999999.0
$consecFail    = 0
$maxConsecFail = 0
$prevLat       = -1.0
$jitterSum     = 0.0
$jitterCount   = 0

for ($i = 1; $i -le $Samples; $i++) {
    $pct = [int](($i / $Samples) * 100)
    Write-Progress -Activity "Capturing network telemetry" `
                   -Status "Sample $i of $Samples to $PingTarget" `
                   -PercentComplete $pct

    $stampNow = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $clockNow = Get-Date -Format "HH:mm:ss"

    $ok = $false
    $ms = 0.0
    try {
        $reply = $pinger.Send($PingTarget, 1000)
        if ($reply -and $reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) {
            $ok = $true
            $ms = [double]$reply.RoundtripTime
        }
    } catch { $ok = $false }

    if ($ok) {
        $consecFail = 0
        $successCount++
        $latSum += $ms
        if ($ms -gt $latMax) { $latMax = $ms }
        if ($ms -lt $latMin) { $latMin = $ms }
        if ($prevLat -ge 0) {
            $d = [math]::Abs($ms - $prevLat)
            $jitterSum += $d
            $jitterCount++
        }
        $prevLat = $ms
        $csvRows.Add("$i,$stampNow,$PingTarget,Success,$(Round2 $ms)")
        $jsRows.Add("{i:$i,t:'$clockNow',ok:1,ms:$(Round2 $ms)}")
    } else {
        $failCount++
        $consecFail++
        if ($consecFail -gt $maxConsecFail) { $maxConsecFail = $consecFail }
        $csvRows.Add("$i,$stampNow,$PingTarget,Failed,")
        $jsRows.Add("{i:$i,t:'$clockNow',ok:0,ms:0}")
    }

    Start-Sleep -Seconds 1
}
Write-Progress -Activity "Capturing network telemetry" -Completed
Set-Content -Path $csvPath -Value ($csvRows -join "`r`n") -Encoding UTF8

$lossPct = Round2 (($failCount / $Samples) * 100)
$avgLat  = 0
$minLat  = 0
$maxLat  = 0
if ($successCount -gt 0) {
    $avgLat = Round2 ($latSum / $successCount)
    $minLat = Round2 $latMin
    $maxLat = Round2 $latMax
}
$jitterVal = 0
if ($jitterCount -gt 0) { $jitterVal = Round2 ($jitterSum / $jitterCount) }

# --------------------------------------------------------- path validation ---

$gwStatus = "unknown"
$gwTest   = "Default gateway was not detected."
if ($routerIP -ne "Not detected") {
    $gwOk = 0
    $gwLines = New-Object System.Collections.Generic.List[string]
    for ($g = 1; $g -le 4; $g++) {
        try {
            $r = $pinger.Send($routerIP, 1000)
            if ($r -and $r.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) {
                $gwOk++
                $gwLines.Add("Reply from ${routerIP}: time=$($r.RoundtripTime)ms")
            } else {
                $gwLines.Add("Request timed out.")
            }
        } catch { $gwLines.Add("Request timed out.") }
    }
    $gwLines.Add("Packets: Sent = 4, Received = $gwOk, Lost = $(4 - $gwOk)")
    $gwTest = ($gwLines -join "`r`n")
    if     ($gwOk -eq 4) { $gwStatus = "clean" }
    elseif ($gwOk -gt 0) { $gwStatus = "lossy" }
    else                 { $gwStatus = "unreachable" }
}

$dnsOk = 0
$dnsTest = ""
try {
    $res = Resolve-DnsName -Name $DnsTarget -Type A -ErrorAction Stop
    $dnsTest = ($res | Format-Table -AutoSize | Out-String)
    if ($res) { $dnsOk = 1 }
} catch {
    $dnsTest = Invoke-Native "nslookup $DnsTarget"
    if ($dnsTest -match 'Address') { $dnsOk = 1 }
}

$httpCode = "000"
$httpTest = ""
try {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $resp = Invoke-WebRequest -Uri $HttpTarget -TimeoutSec 15 -UseBasicParsing -ErrorAction Stop
    $sw.Stop()
    $httpCode = "$([int]$resp.StatusCode)"
    $httpTest = "target=$HttpTarget`r`nstatus=$httpCode`r`ntotal=$([math]::Round($sw.Elapsed.TotalSeconds,3))s`r`nbytes=$($resp.RawContentLength)"
} catch {
    $httpTest = "target=$HttpTarget`r`nstatus=failed`r`nerror=$($_.Exception.Message)"
    if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
        $httpCode = "$([int]$_.Exception.Response.StatusCode)"
    }
}

$traceTest = Invoke-Native "tracert -d -w 600 -h 12 $PingTarget"
$hop1Loss = 0
$traceLines = @($traceTest -split "`r?`n" | Where-Object { $_ -match '^\s*\d+\s' })
if ($traceLines.Count -gt 0 -and $traceLines[0] -match '\*\s+\*\s+\*') { $hop1Loss = 1 }

$routeTable = Invoke-Native "route print -4"

Write-Log "Loss $lossPct pct  Avg $avgLat ms  Jitter $jitterVal ms  MaxDrop $maxConsecFail  Gateway $gwStatus  DNS $dnsOk  HTTP $httpCode"

# ------------------------------------------------------------ data payload ---

$epochMs = [long]([DateTimeOffset]::Now.ToUnixTimeMilliseconds())
$sigPctJs = if ($null -eq $sigPct) { "null" } else { "$sigPct" }

$sb = New-Object System.Text.StringBuilder
[void]$sb.Append("{")
[void]$sb.Append("v:'$TOOL_VERSION',")
[void]$sb.Append("plat:'Windows',")
[void]$sb.Append("id:'${stamp}_${safeLabel}',")
[void]$sb.Append("ts:$epochMs,")
[void]$sb.Append("loc:$(ConvertTo-JsString $Location),")
[void]$sb.Append("host:$(ConvertTo-JsString $hostName),")
[void]$sb.Append("user:$(ConvertTo-JsString $consoleUser),")
[void]$sb.Append("model:$(ConvertTo-JsString $modelName),")
[void]$sb.Append("modelId:$(ConvertTo-JsString $modelID),")
[void]$sb.Append("chip:$(ConvertTo-JsString $chipName),")
[void]$sb.Append("serial:$(ConvertTo-JsString $serialNum),")
[void]$sb.Append("osName:$(ConvertTo-JsString $osName),")
[void]$sb.Append("osVer:$(ConvertTo-JsString $osVer),")
[void]$sb.Append("osBuild:$(ConvertTo-JsString $osBuild),")
[void]$sb.Append("wifiDev:$(ConvertTo-JsString $wifiDev),")
[void]$sb.Append("iface:$(ConvertTo-JsString $ifaceNm),")
[void]$sb.Append("link:$(ConvertTo-JsString $linkType),")
[void]$sb.Append("ip:$(ConvertTo-JsString $activeIP),")
[void]$sb.Append("router:$(ConvertTo-JsString $routerIP),")
[void]$sb.Append("dns:$(ConvertTo-JsString $dnsServers),")
[void]$sb.Append("ssid:$(ConvertTo-JsString $ssidName),")
[void]$sb.Append("bssid:$(ConvertTo-JsString $bssidVal),")
[void]$sb.Append("phy:$(ConvertTo-JsString $phyMode),")
[void]$sb.Append("chan:$(ConvertTo-JsString $wChannel),")
[void]$sb.Append("sec:$(ConvertTo-JsString $wSecurity),")
[void]$sb.Append("tx:$(ConvertTo-JsString $txRate),")
[void]$sb.Append("haveSig:$($haveSig.ToString().ToLower()),")
[void]$sb.Append("haveNoise:false,")
[void]$sb.Append("sigPct:$sigPctJs,")
[void]$sb.Append("rssi:$rssiVal,")
[void]$sb.Append("noise:0,")
[void]$sb.Append("snr:0,")
[void]$sb.Append("nbTotal:$neighborTotal,")
[void]$sb.Append("nb:$nbJs,")
[void]$sb.Append("target:$(ConvertTo-JsString $PingTarget),")
[void]$sb.Append("n:$Samples,")
[void]$sb.Append("ok:$successCount,")
[void]$sb.Append("bad:$failCount,")
[void]$sb.Append("loss:$lossPct,")
[void]$sb.Append("avg:$avgLat,")
[void]$sb.Append("min:$minLat,")
[void]$sb.Append("max:$maxLat,")
[void]$sb.Append("jit:$jitterVal,")
[void]$sb.Append("drop:$maxConsecFail,")
[void]$sb.Append("gw:'$gwStatus',")
[void]$sb.Append("dnsOk:$dnsOk,")
[void]$sb.Append("httpCode:$(ConvertTo-JsString $httpCode),")
[void]$sb.Append("hop1:$hop1Loss,")
[void]$sb.Append("samples:[" + ($jsRows -join ",") + "],")
[void]$sb.Append("raw:{")
[void]$sb.Append("gw:$(ConvertTo-JsString $gwTest),")
[void]$sb.Append("dns:$(ConvertTo-JsString $dnsTest),")
[void]$sb.Append("http:$(ConvertTo-JsString $httpTest),")
[void]$sb.Append("trace:$(ConvertTo-JsString $traceTest),")
[void]$sb.Append("route:$(ConvertTo-JsString $routeTable),")
[void]$sb.Append("ifc:$(ConvertTo-JsString $wlanIfaceRaw),")
[void]$sb.Append("scutil:$(ConvertTo-JsString $ipconfigRaw),")
[void]$sb.Append("pref:$(ConvertTo-JsString $wlanProfileRaw),")
[void]$sb.Append("air:$(ConvertTo-JsString $wlanDriverRaw)},")
[void]$sb.Append("out:$(ConvertTo-JsString $outDir)")
[void]$sb.Append("}")
$capJs = $sb.ToString()

# ------------------------------------------------------------- stylesheet ---

$cssText = @'
*,*::before,*::after{box-sizing:border-box}
:root{--bg:#05080d;--bg2:#080d15;--grid-line:rgba(0,229,255,.045);--panel:rgba(11,17,27,.86);--panel-solid:#0b111b;--border:rgba(0,229,255,.16);--border-soft:rgba(255,255,255,.07);--txt:#d7e6f2;--txt-dim:#7d93a8;--txt-mute:#4e6275;--cyan:#00e5ff;--green:#00e08a;--amber:#ffb020;--red:#ff3b57;--violet:#a97bff;--mono:'SFMono-Regular',ui-monospace,Menlo,Consolas,monospace;--sans:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;--r:10px;--shadow:0 10px 34px rgba(0,0,0,.6)}
html[data-theme='light']{--bg:#eef2f7;--bg2:#e4eaf2;--grid-line:rgba(0,80,120,.05);--panel:rgba(255,255,255,.92);--panel-solid:#fff;--border:rgba(0,110,150,.22);--border-soft:rgba(0,0,0,.09);--txt:#0f1c28;--txt-dim:#456173;--txt-mute:#7b93a5;--cyan:#0092b8;--green:#00875a;--amber:#b06f00;--red:#c62a42;--violet:#6b3fd4;--shadow:0 8px 26px rgba(20,40,60,.14)}
html,body{height:100%}
body{margin:0;background:var(--bg);color:var(--txt);font-family:var(--sans);font-size:14px;line-height:1.5;-webkit-font-smoothing:antialiased;overflow-x:hidden;background-image:linear-gradient(var(--grid-line) 1px,transparent 1px),linear-gradient(90deg,var(--grid-line) 1px,transparent 1px);background-size:46px 46px}
body::before{content:'';position:fixed;inset:0;pointer-events:none;z-index:0;background:radial-gradient(1100px 620px at 12% -8%,rgba(0,229,255,.10),transparent 62%),radial-gradient(900px 520px at 92% 4%,rgba(169,123,255,.08),transparent 60%)}
::-webkit-scrollbar{width:9px;height:9px}
::-webkit-scrollbar-track{background:transparent}
::-webkit-scrollbar-thumb{background:rgba(0,229,255,.22);border-radius:9px}
::-webkit-scrollbar-thumb:hover{background:rgba(0,229,255,.4)}
button,input,select{font-family:inherit;color:inherit}
button{cursor:pointer}
a{color:var(--cyan)}
.topbar{position:sticky;top:0;z-index:60;display:flex;align-items:center;gap:10px;flex-wrap:wrap;padding:8px 14px;background:linear-gradient(180deg,rgba(5,8,13,.97),rgba(5,8,13,.86));border-bottom:1px solid var(--border);backdrop-filter:blur(14px);-webkit-backdrop-filter:blur(14px)}
html[data-theme='light'] .topbar{background:linear-gradient(180deg,rgba(255,255,255,.97),rgba(255,255,255,.86))}
.brand{display:flex;align-items:center;gap:8px;font-family:var(--mono);font-weight:700;letter-spacing:.13em;font-size:13px;white-space:nowrap}
.brand .bolt{font-size:17px;filter:drop-shadow(0 0 8px var(--cyan))}
.brand .b1{color:var(--cyan)}
.brand .b2{color:var(--txt)}
.ver{font-size:9px;color:var(--txt-mute);border:1px solid var(--border-soft);padding:1px 5px;border-radius:4px;font-family:var(--mono)}
.live{display:inline-flex;align-items:center;gap:5px;font-family:var(--mono);font-size:10px;letter-spacing:.14em;color:var(--green);border:1px solid rgba(0,224,138,.34);background:rgba(0,224,138,.08);padding:2px 7px;border-radius:20px}
.live .dot{width:6px;height:6px;border-radius:50%;background:var(--green);animation:blip 1.6s infinite}
@keyframes blip{0%,100%{opacity:1;box-shadow:0 0 0 0 rgba(0,224,138,.6)}55%{opacity:.35;box-shadow:0 0 0 7px rgba(0,224,138,0)}}
.tb-spacer{flex:1 1 auto}
.clockbox{font-family:var(--mono);font-size:11px;color:var(--txt-dim);display:flex;flex-direction:column;line-height:1.28;text-align:right}
.clockbox b{color:var(--txt);font-size:13px;letter-spacing:.05em}
.tb-btn{background:rgba(255,255,255,.035);border:1px solid var(--border-soft);color:var(--txt-dim);border-radius:7px;padding:5px 9px;font-size:11px;font-family:var(--mono);letter-spacing:.06em;display:inline-flex;align-items:center;gap:5px;transition:.16s;white-space:nowrap}
.tb-btn:hover{border-color:var(--cyan);color:var(--cyan);background:rgba(0,229,255,.09);transform:translateY(-1px)}
.kbd{font-family:var(--mono);font-size:9px;border:1px solid var(--border-soft);border-radius:3px;padding:0 4px;color:var(--txt-mute);background:rgba(255,255,255,.04)}
.searchbtn{min-width:160px;justify-content:space-between}
.subbar{position:sticky;top:46px;z-index:55;display:flex;align-items:center;gap:6px;overflow-x:auto;padding:7px 14px;background:rgba(5,8,13,.9);border-bottom:1px solid var(--border-soft);backdrop-filter:blur(10px);scrollbar-width:none}
html[data-theme='light'] .subbar{background:rgba(255,255,255,.9)}
.subbar::-webkit-scrollbar{display:none}
.sb-label{font-family:var(--mono);font-size:9px;letter-spacing:.16em;color:var(--txt-mute);white-space:nowrap;padding-right:2px}
.chip{font-family:var(--mono);font-size:10px;letter-spacing:.05em;white-space:nowrap;border:1px solid var(--border-soft);background:rgba(255,255,255,.03);color:var(--txt-mute);padding:3px 9px;border-radius:20px;transition:.16s}
.chip:hover{color:var(--txt);border-color:var(--border)}
.chip.on{color:var(--cyan);border-color:rgba(0,229,255,.45);background:rgba(0,229,255,.12);box-shadow:0 0 12px rgba(0,229,255,.14) inset}
.range-pills{display:flex;gap:3px;margin-left:auto;align-items:center}
.pill{font-family:var(--mono);font-size:10px;padding:3px 8px;border-radius:5px;border:1px solid var(--border-soft);background:transparent;color:var(--txt-mute)}
.pill.on{background:rgba(0,229,255,.16);border-color:var(--cyan);color:var(--cyan)}
.wrap{position:relative;z-index:1;padding:14px;max-width:1680px;margin:0 auto}
#alertBanner{display:none;align-items:center;gap:10px;margin-bottom:12px;padding:11px 15px;border-radius:var(--r);border:1px solid rgba(255,59,87,.5);background:linear-gradient(90deg,rgba(255,59,87,.2),rgba(255,59,87,.05));font-family:var(--mono);font-size:12px;letter-spacing:.06em;color:#ffd7dd;animation:pulseBar 2.2s infinite}
@keyframes pulseBar{0%,100%{box-shadow:0 0 0 0 rgba(255,59,87,.35)}50%{box-shadow:0 0 22px 2px rgba(255,59,87,.14)}}
#alertBanner.show{display:flex}
#alertBanner .x{margin-left:auto;background:none;border:none;color:#ffd7dd;font-size:14px}
.hero{border:1px solid var(--border);border-radius:14px;background:var(--panel);backdrop-filter:blur(12px);box-shadow:var(--shadow);overflow:hidden;margin-bottom:14px;position:relative}
.hero::after{content:'';position:absolute;left:0;right:0;top:0;height:2px;background:linear-gradient(90deg,transparent,var(--cyan),var(--violet),transparent);background-size:200% 100%;animation:shimmer 4.5s linear infinite}
@keyframes shimmer{0%{background-position:200% 0}100%{background-position:-200% 0}}
.hero-top{display:flex;flex-wrap:wrap;align-items:center;gap:14px;padding:15px 18px;border-bottom:1px solid var(--border-soft)}
.statuslamp{width:62px;height:62px;border-radius:50%;display:grid;place-items:center;font-size:27px;flex:none;position:relative}
.statuslamp::after{content:'';position:absolute;inset:-6px;border-radius:50%;border:2px solid currentColor;opacity:.32;animation:ring 2.4s infinite}
@keyframes ring{0%{transform:scale(.86);opacity:.55}100%{transform:scale(1.22);opacity:0}}
.lamp-ok{background:radial-gradient(circle at 35% 30%,#00e08a,#046b45);color:var(--green);box-shadow:0 0 34px rgba(0,224,138,.5)}
.lamp-warn{background:radial-gradient(circle at 35% 30%,#ffc857,#8a5b00);color:var(--amber);box-shadow:0 0 34px rgba(255,176,32,.5)}
.lamp-bad{background:radial-gradient(circle at 35% 30%,#ff6076,#8d0d22);color:var(--red);box-shadow:0 0 34px rgba(255,59,87,.5)}
.hero-id{flex:1 1 250px;min-width:210px}
.hero-id .lbl{font-family:var(--mono);font-size:9.5px;letter-spacing:.2em;color:var(--txt-mute)}
.hero-id h1{margin:2px 0 3px;font-size:25px;font-weight:750;letter-spacing:-.4px}
.hero-id .sub{font-family:var(--mono);font-size:11px;color:var(--txt-dim)}
.hero-rate{text-align:right;flex:none}
.hero-rate .big{font-family:var(--mono);font-size:40px;font-weight:750;line-height:1;letter-spacing:-1.5px}
.hero-rate .unit{font-family:var(--mono);font-size:10px;color:var(--txt-mute);letter-spacing:.16em}
.statuspill{display:inline-flex;align-items:center;gap:5px;font-family:var(--mono);font-size:10px;letter-spacing:.14em;padding:3px 9px;border-radius:20px;border:1px solid}
.sp-ok{color:var(--green);border-color:rgba(0,224,138,.45);background:rgba(0,224,138,.1)}
.sp-warn{color:var(--amber);border-color:rgba(255,176,32,.45);background:rgba(255,176,32,.1)}
.sp-bad{color:var(--red);border-color:rgba(255,59,87,.45);background:rgba(255,59,87,.1)}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(148px,1fr));gap:1px;background:var(--border-soft)}
.kpi{background:var(--panel-solid);padding:11px 14px;transition:.18s}
.kpi:hover{background:rgba(0,229,255,.06);transform:translateY(-2px)}
.kpi .k{font-family:var(--mono);font-size:9px;letter-spacing:.15em;color:var(--txt-mute);margin-bottom:3px}
.kpi .v{font-family:var(--mono);font-size:19px;font-weight:700;letter-spacing:-.5px}
.kpi .d{font-family:var(--mono);font-size:9.5px;color:var(--txt-dim);margin-top:1px}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(360px,1fr));gap:13px;align-items:start}
.panel{border:1px solid var(--border-soft);border-radius:var(--r);background:var(--panel);backdrop-filter:blur(10px);box-shadow:var(--shadow);overflow:hidden;animation:fadeUp .42s cubic-bezier(.2,.7,.3,1) both;position:relative}
@keyframes fadeUp{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:none}}
.panel.wide{grid-column:span 2}
.panel.glow{border-color:rgba(0,229,255,.5);box-shadow:0 0 26px rgba(0,229,255,.16)}
.p-head{display:flex;align-items:center;gap:8px;padding:9px 12px;border-bottom:1px solid var(--border-soft);background:linear-gradient(180deg,rgba(0,229,255,.05),transparent);cursor:pointer;user-select:none}
.p-head h3{margin:0;font-family:var(--mono);font-size:11px;letter-spacing:.14em;font-weight:600;color:var(--txt);display:flex;align-items:center;gap:6px;flex:1;min-width:0}
.badge{font-family:var(--mono);font-size:9px;padding:1px 6px;border-radius:20px;border:1px solid var(--border-soft);color:var(--txt-mute)}
.badge.new{color:var(--cyan);border-color:rgba(0,229,255,.4);background:rgba(0,229,255,.12)}
.badge.ok{color:var(--green);border-color:rgba(0,224,138,.4);background:rgba(0,224,138,.12)}
.badge.bad{color:var(--red);border-color:rgba(255,59,87,.4);background:rgba(255,59,87,.12)}
.p-act{background:none;border:none;color:var(--txt-mute);font-size:12px;padding:1px 4px;border-radius:4px;line-height:1}
.p-act:hover{color:var(--cyan);background:rgba(0,229,255,.1)}
.p-body{padding:12px;max-height:2400px;overflow:auto;transition:max-height .3s ease,padding .3s ease}
.panel.collapsed .p-body{max-height:0;padding-top:0;padding-bottom:0;overflow:hidden}
.panel.collapsed .caret{transform:rotate(-90deg)}
.caret{transition:transform .25s;display:inline-block}
.mono{font-family:var(--mono)}
.dim{color:var(--txt-dim)}
.mute{color:var(--txt-mute)}
.c-green{color:var(--green)}
.c-amber{color:var(--amber)}
.c-red{color:var(--red)}
.c-cyan{color:var(--cyan)}
.c-violet{color:var(--violet)}
.hint{font-size:11px;color:var(--txt-mute);font-family:var(--mono);line-height:1.5}
.row{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.btn{background:rgba(0,229,255,.08);border:1px solid var(--border);color:var(--cyan);border-radius:7px;padding:7px 12px;font-family:var(--mono);font-size:11px;letter-spacing:.07em;transition:.16s}
.btn:hover{background:rgba(0,229,255,.2);transform:translateY(-1px);box-shadow:0 4px 14px rgba(0,229,255,.16)}
.btn.ghost{background:transparent;border-color:var(--border-soft);color:var(--txt-dim)}
.btn.ghost:hover{color:var(--txt);border-color:var(--border);box-shadow:none}
.btn.sm{padding:4px 8px;font-size:10px}
table{width:100%;border-collapse:collapse;font-family:var(--mono);font-size:11px}
th{text-align:left;color:var(--txt-mute);font-weight:600;letter-spacing:.1em;font-size:9px;padding:5px 6px;border-bottom:1px solid var(--border-soft);text-transform:uppercase}
td{padding:6px;border-bottom:1px solid rgba(255,255,255,.04)}
tbody tr{transition:.14s}
tbody tr:hover{background:rgba(0,229,255,.06)}
tbody tr.best td{background:rgba(0,224,138,.09)}
tbody tr.bad td{background:rgba(255,59,87,.08)}
canvas{width:100%;display:block;border-radius:7px}
.empty{text-align:center;padding:20px 8px;color:var(--txt-mute);font-family:var(--mono);font-size:11px}
.result{margin-top:10px;border:1px solid var(--border);border-radius:9px;padding:11px;background:rgba(0,229,255,.05)}
.result .line{display:flex;justify-content:space-between;font-family:var(--mono);font-size:11px;padding:2.5px 0;border-bottom:1px dashed rgba(255,255,255,.06);gap:10px}
.result .line:last-child{border:none}
.advice{margin-top:9px;padding:9px 10px;border-radius:7px;font-size:12px;line-height:1.55;border-left:3px solid}
.advice.good{background:rgba(0,224,138,.1);border-color:var(--green)}
.advice.warn{background:rgba(255,176,32,.1);border-color:var(--amber)}
.advice.bad{background:rgba(255,59,87,.1);border-color:var(--red)}
.rc{margin-bottom:9px}
.rc .rc-top{display:flex;align-items:baseline;gap:8px;font-family:var(--mono);font-size:11.5px}
.rc .rc-top b{flex:1}
.rc .bar{height:7px;border-radius:6px;background:rgba(255,255,255,.08);overflow:hidden;margin:4px 0 3px}
.rc .bar i{display:block;height:100%;transition:width .5s}
.rc .ev{font-size:11px;color:var(--txt-dim);line-height:1.5}
.xp{border-left:3px solid var(--cyan);background:rgba(0,229,255,.045);border-radius:8px;padding:12px 14px;margin-bottom:10px}
.xp.ok{border-left-color:var(--green);background:rgba(0,224,138,.05)}
.xp.warn{border-left-color:var(--amber);background:rgba(255,176,32,.05)}
.xp.bad{border-left-color:var(--red);background:rgba(255,59,87,.05)}
.xp h4{margin:0 0 6px;font-size:13.5px;color:var(--txt);font-weight:650}
.xp .an{font-style:italic;color:var(--txt-dim);font-size:12.5px;line-height:1.55;margin-bottom:8px}
.xp .rs{font-family:var(--mono);font-size:11px;letter-spacing:.06em;color:var(--cyan);margin-bottom:6px}
.xp.ok .rs{color:var(--green)}
.xp.warn .rs{color:var(--amber)}
.xp.bad .rs{color:var(--red)}
.xp .mn{font-size:12.5px;line-height:1.62;color:var(--txt)}
.vbox{border:1px solid var(--border);border-radius:11px;padding:16px 17px;margin-bottom:15px;background:rgba(0,229,255,.06)}
.vbox.ok{border-color:rgba(0,224,138,.5);background:rgba(0,224,138,.07)}
.vbox.warn{border-color:rgba(255,176,32,.5);background:rgba(255,176,32,.07)}
.vbox.bad{border-color:rgba(255,59,87,.5);background:rgba(255,59,87,.07)}
.vbox .vt{font-size:18px;font-weight:720;margin-bottom:7px;letter-spacing:-.2px}
.vbox .vs{font-size:13.5px;line-height:1.65}
.msgbox{border:1px dashed var(--border);border-radius:9px;padding:13px 15px;margin-top:10px;background:rgba(0,0,0,.24);font-size:12.5px;line-height:1.7;white-space:pre-wrap}
html[data-theme='light'] .msgbox{background:rgba(0,0,0,.04)}
.glos{display:grid;grid-template-columns:132px 1fr;gap:5px 12px;font-size:12.5px;line-height:1.55;margin-top:4px}
.glos b{color:var(--cyan);font-family:var(--mono);font-size:11px;letter-spacing:.04em;padding-top:2px}
@media(max-width:600px){.glos{grid-template-columns:1fr;gap:2px 0}.glos b{padding-top:8px}}
pre{white-space:pre-wrap;word-break:break-word;background:rgba(0,0,0,.34);border:1px solid var(--border-soft);border-radius:8px;padding:11px;margin:0;font-family:var(--mono);font-size:11px;color:var(--txt-dim);max-height:340px;overflow:auto}
html[data-theme='light'] pre{background:rgba(0,0,0,.04)}
.overlay{position:fixed;inset:0;z-index:200;background:rgba(2,5,10,.8);backdrop-filter:blur(7px);display:none;align-items:flex-start;justify-content:center;padding:9vh 16px}
.overlay.show{display:flex}
.cmdbox{width:100%;max-width:620px;background:var(--panel-solid);border:1px solid var(--border);border-radius:13px;box-shadow:0 24px 70px rgba(0,0,0,.7);overflow:hidden;animation:pop .2s cubic-bezier(.2,.8,.3,1)}
@keyframes pop{from{transform:translateY(-14px) scale(.98);opacity:0}to{transform:none;opacity:1}}
.cmdbox input{width:100%;border:none;background:transparent;padding:15px 17px;font-size:15px;outline:none}
.cmd-list{max-height:52vh;overflow:auto;border-top:1px solid var(--border-soft)}
.cmd-item{display:flex;align-items:center;gap:10px;padding:9px 16px;font-size:13px;cursor:pointer;border-left:3px solid transparent}
.cmd-item .cs{margin-left:auto;font-family:var(--mono);font-size:9px;color:var(--txt-mute);letter-spacing:.1em}
.cmd-item.sel{background:rgba(0,229,255,.13);border-left-color:var(--cyan)}
.cmd-foot{display:flex;gap:12px;padding:7px 16px;border-top:1px solid var(--border-soft);font-family:var(--mono);font-size:9px;color:var(--txt-mute);letter-spacing:.08em}
.modal{width:100%;max-width:660px;background:var(--panel-solid);border:1px solid var(--border);border-radius:13px;box-shadow:0 24px 70px rgba(0,0,0,.7);overflow:hidden;animation:pop .2s;max-height:82vh;display:flex;flex-direction:column}
.modal h2{margin:0;padding:13px 17px;font-family:var(--mono);font-size:12px;letter-spacing:.16em;border-bottom:1px solid var(--border-soft);display:flex;align-items:center;gap:8px}
.modal h2 .x{margin-left:auto;background:none;border:none;color:var(--txt-mute);font-size:16px}
.modal-body{padding:15px 17px;overflow:auto}
#toasts{position:fixed;bottom:16px;right:16px;z-index:300;display:flex;flex-direction:column;gap:8px;max-width:340px}
.toast{border:1px solid var(--border);background:var(--panel-solid);border-radius:9px;padding:10px 13px;font-family:var(--mono);font-size:11.5px;box-shadow:var(--shadow);border-left:3px solid var(--cyan);animation:slideIn .25s}
@keyframes slideIn{from{transform:translateX(40px);opacity:0}to{transform:none;opacity:1}}
.toast.ok{border-left-color:var(--green)}
.toast.warn{border-left-color:var(--amber)}
.toast.err{border-left-color:var(--red)}
footer{max-width:1680px;margin:20px auto 0;padding:22px 14px 34px;font-family:var(--mono);font-size:10px;color:var(--txt-mute);line-height:1.75;border-top:1px solid var(--border-soft);position:relative;z-index:1}
@media(max-width:820px){.grid{grid-template-columns:1fr}.panel.wide{grid-column:span 1}.hero-rate{text-align:left}.hero-rate .big{font-size:32px}.hero-id h1{font-size:20px}.wrap{padding:10px}}
@media print{body{background:#fff;color:#000}.topbar,.subbar,.overlay,#toasts,.p-act,.noprint{display:none!important}.panel{break-inside:avoid;border:1px solid #999}.p-body{max-height:none!important}}
'@

# ------------------------------------------------------------ page shell ---

$shellText = @'
</style></head><body>
<div class='topbar' id='topbar'>
<div class='brand'><span class='bolt'>&#128225;</span><span class='b1'>WIFI</span><span class='b2'>CONNECTIVITY MONITOR</span><span class='ver'>v3.1</span></div>
<span class='live'><span class='dot'></span>CAPTURED</span>
<span class='statuspill sp-ok' id='platPill'>PLATFORM</span>
<span class='statuspill sp-ok' id='locPill'>LOCATION</span>
<span class='statuspill sp-ok' id='linkPill'>LINK</span>
<div class='tb-spacer'></div>
<button class='tb-btn' data-act='explain' title='Plain English briefing (X)'>&#128172; EXPLAIN</button>
<button class='tb-btn searchbtn' data-act='palette'>&#128269; Search / Commands <span class='kbd'>&#8984;K</span></button>
<button class='tb-btn' data-act='theme' title='Toggle theme (T)'>&#9680;</button>
<button class='tb-btn' data-act='rerender' title='Re-render panels (R)'>&#8635;</button>
<button class='tb-btn' data-act='export' title='Export snapshot (E)'>&#10515; EXPORT</button>
<button class='tb-btn' data-act='print' title='Print / PDF (P)'>&#128424; PDF</button>
<button class='tb-btn' data-act='full' title='Fullscreen (F)'>&#9974;</button>
<div class='clockbox'><b id='clockLocal'>--:--:--</b><span id='clockMeta'>CAPTURE</span></div>
</div>
<div class='subbar' id='subbar'>
<span class='sb-label'>PANELS</span>
<div id='chipRow' style='display:flex;gap:6px'></div>
<div class='range-pills' id='rangePills'>
<span class='sb-label'>SAMPLES</span>
<button class='pill' data-win='25'>25</button>
<button class='pill' data-win='50'>50</button>
<button class='pill' data-win='100'>100</button>
<button class='pill on' data-win='0'>ALL</button>
</div>
</div>
<div class='wrap'>
<div id='alertBanner'><span style='font-size:16px'>&#9888;</span><span id='alertText'></span><button class='x' data-act='dismiss'>&#10005;</button></div>
<div class='hero'>
<div class='hero-top'>
<div class='statuslamp lamp-ok' id='lamp'>&#128994;</div>
<div class='hero-id'>
<div class='lbl'>CAPTURE VERDICT</div>
<h1 id='heroTitle'>ANALYZING</h1>
<div class='sub' id='heroSub'></div>
</div>
<div class='hero-rate'>
<div class='big' id='heroBig'>0%</div>
<div class='unit'>PACKET LOSS TO 8.8.8.8</div>
<div style='margin-top:6px'><span class='statuspill sp-ok' id='heroPill'>--</span></div>
</div>
</div>
<div class='kpis' id='kpis'></div>
</div>
<div class='grid' id='grid'></div>
</div>
<footer>
<div><b class='c-cyan'>WIFI CONNECTIVITY MONITOR v3.1</b> &middot; IT Support &middot; read-only capture, no network settings were changed.</div>
<div>New in v3.1: a <b>Plain English Briefing</b> panel that explains every measurement in everyday language with analogies, plus a ready-to-send status message. Share this report with non-technical colleagues.</div>
<div>Root cause scores are weighted heuristics derived from the measurements in this capture. They rank where to investigate first; they are not a substitute for switch, controller or ISP-side telemetry.</div>
<div id='platNote'></div>
<div class='noprint'>100% local &middot; nothing leaves this machine &middot; capture ledger stored in this browser profile localStorage &middot; <span class='kbd'>&#8984;K</span> commands &middot; <span class='kbd'>?</span> shortcuts.</div>
</footer>
<div class='overlay' id='palette'><div class='cmdbox'><input id='cmdInput' placeholder='Jump to a panel or run a command...' autocomplete='off' spellcheck='false'><div class='cmd-list' id='cmdList'></div><div class='cmd-foot'><span>UP/DOWN NAVIGATE</span><span>ENTER RUN</span><span>ESC CLOSE</span></div></div></div>
<div class='overlay' id='helpModal'><div class='modal'><h2>&#9000; KEYBOARD SHORTCUTS <button class='x' data-act='closeHelp'>&#10005;</button></h2><div class='modal-body'><table><tbody>
<tr><td><span class='kbd'>&#8984;K</span> / <span class='kbd'>Ctrl K</span></td><td class='dim'>Open command palette</td></tr>
<tr><td><span class='kbd'>X</span></td><td class='dim'>Jump to the Plain English Briefing</td></tr>
<tr><td><span class='kbd'>T</span></td><td class='dim'>Toggle dark / light theme</td></tr>
<tr><td><span class='kbd'>F</span></td><td class='dim'>Toggle fullscreen</td></tr>
<tr><td><span class='kbd'>R</span></td><td class='dim'>Re-render all panels</td></tr>
<tr><td><span class='kbd'>E</span></td><td class='dim'>Export capture snapshot (JSON)</td></tr>
<tr><td><span class='kbd'>P</span></td><td class='dim'>Print / save as PDF</td></tr>
<tr><td><span class='kbd'>C</span></td><td class='dim'>Copy the Copilot root cause prompt</td></tr>
<tr><td><span class='kbd'>M</span></td><td class='dim'>Copy the ready-to-send status message</td></tr>
<tr><td><span class='kbd'>1..4</span></td><td class='dim'>Sample window (25 / 50 / 100 / ALL)</td></tr>
<tr><td><span class='kbd'>?</span></td><td class='dim'>This help</td></tr>
<tr><td><span class='kbd'>ESC</span></td><td class='dim'>Close any overlay</td></tr>
</tbody></table></div></div></div>
<div id='toasts'></div>
<script>
'@

# ------------------------------------------------------ dashboard engine ---

$jsText = @'
'use strict';
var LS={get:function(k,d){try{var v=localStorage.getItem('wifimon3.'+k);return v===null?d:JSON.parse(v);}catch(e){return d;}},set:function(k,v){try{localStorage.setItem('wifimon3.'+k,JSON.stringify(v));}catch(e){}}};
function esc(s){if(s===null||s===undefined)return '';s=String(s);return s.split('&').join('&amp;').split('<').join('&lt;').split('>').join('&gt;').split(String.fromCharCode(34)).join('&quot;').split(String.fromCharCode(39)).join('&#39;');}
function q(s){return document.querySelector(s);}
function chanNum(t){var m=String(t).match(/[0-9]+/);return m?parseInt(m[0],10):0;}
function bandOf(c){if(c<=0)return 'unknown';if(c<=14)return '2.4 GHz';if(c<=177)return '5 GHz';return '6 GHz';}
function fx(v,n){return (Math.round(v*Math.pow(10,n))/Math.pow(10,n)).toFixed(n);}
var WIN=LS.get('win',0);
var CHAN=chanNum(CAP.chan),BAND=bandOf(CHAN),ISWIFI=(CAP.link.indexOf('Wi-Fi')===0);
var HAVENOISE=(CAP.haveNoise===undefined)?CAP.haveSig:CAP.haveNoise;
var PLAT=CAP.plat||'macOS';
var HASPCT=(CAP.sigPct!==undefined&&CAP.sigPct!==null);
var Ledger={all:function(){return LS.get('caps',[]);},
save:function(){var a=this.all();a=a.filter(function(x){return x.id!==CAP.id;});a.push({id:CAP.id,ts:CAP.ts,loc:CAP.loc,link:CAP.link,ssid:CAP.ssid,chan:CAP.chan,rssi:CAP.rssi,snr:CAP.snr,haveSig:CAP.haveSig,haveNoise:CAP.haveNoise,plat:CAP.plat,loss:CAP.loss,avg:CAP.avg,max:CAP.max,jit:CAP.jit,drop:CAP.drop,nbTotal:CAP.nbTotal,gw:CAP.gw,n:CAP.n});a.sort(function(x,y){return x.ts-y.ts;});if(a.length>14)a=a.slice(a.length-14);LS.set('caps',a);return a;},
clear:function(){LS.set('caps',[]);}};
var CAPS=Ledger.save();
function toast(m,kind,ms){var t=document.createElement('div');t.className='toast '+(kind||'');t.textContent=m;q('#toasts').appendChild(t);setTimeout(function(){t.style.transition='opacity .3s,transform .3s';t.style.opacity='0';t.style.transform='translateX(40px)';setTimeout(function(){t.remove();},320);},ms||4200);}
var Score={run:function(){
var s=[],ev,sc;
sc=0;ev=[];
if(HAVENOISE&&CAP.snr>0&&CAP.snr<25){sc+=45;ev.push('SNR is '+CAP.snr+' dB. Healthy is 25 dB or better, so the noise floor here is elevated.');}
else if(HAVENOISE&&CAP.snr>=25&&CAP.snr<32){sc+=18;ev.push('SNR '+CAP.snr+' dB is workable but not comfortable.');}
else if(!HAVENOISE&&ISWIFI){ev.push('This platform does not report a noise floor, so SNR could not be measured. Interference is scored from band, channel density and jitter only.');}
if(CAP.nbTotal>12){sc+=25;ev.push(CAP.nbTotal+' neighbouring networks detected. This is a dense multi-tenant RF environment.');}
else if(CAP.nbTotal>6){sc+=12;ev.push(CAP.nbTotal+' neighbouring networks detected.');}
if(BAND==='2.4 GHz'){sc+=18;ev.push('Client is associated on 2.4 GHz (channel '+CHAN+'), the most congested and interference-prone band.');}
if(ISWIFI&&CAP.jit>25){sc+=12;ev.push('Jitter of '+CAP.jit+' ms suggests contention for airtime.');}
if(!ISWIFI){sc=Math.round(sc*0.15);ev=['This capture ran over a wired link, so wireless interference is largely excluded here.'];}
s.push({k:'RF Interference / Noise',sc:sc,ev:ev});
sc=0;ev=[];
if(CAP.haveSig&&CAP.rssi<=-78){sc+=55;ev.push('RSSI '+CAP.rssi+' dBm is very weak. The client is at the edge of usable coverage.');}
else if(CAP.haveSig&&CAP.rssi<=-72){sc+=35;ev.push('RSSI '+CAP.rssi+' dBm is marginal for this position.');}
else if(CAP.haveSig&&CAP.rssi<=-67){sc+=15;ev.push('RSSI '+CAP.rssi+' dBm is acceptable but not strong.');}
if(CAP.haveSig&&CAP.rssi<=-72&&CAP.loss>0){sc+=20;ev.push('Weak signal is coinciding with measured packet loss.');}
if(!ISWIFI){sc=0;ev=['Wired capture. Access point coverage is not a factor for this measurement.'];}
s.push({k:'Access Point Coverage / Placement',sc:sc,ev:ev});
sc=0;ev=[];
if(CAP.drop>=5){sc+=60;ev.push('Longest unbroken outage was '+CAP.drop+' consecutive samples. That is a genuine link drop, not latency.');}
else if(CAP.drop>=3){sc+=40;ev.push(CAP.drop+' consecutive samples failed, consistent with a roaming event or association loss.');}
else if(CAP.drop===2){sc+=18;ev.push('Two consecutive samples failed, so a brief interruption was observed.');}
if(ISWIFI&&CAP.drop>=3&&CAP.haveSig&&CAP.rssi>-70){sc+=15;ev.push('Drops are occurring despite adequate signal. Look at roaming thresholds and access point overlap.');}
s.push({k:'Roaming / Association Loss',sc:sc,ev:ev});
sc=0;ev=[];
if(CAP.gw==='unreachable'){sc+=70;ev.push('The default gateway did not respond at all.');}
else if(CAP.gw==='lossy'){sc+=45;ev.push('The default gateway responded but with packet loss, so the problem starts inside the office.');}
else if(CAP.gw==='unknown'){sc+=25;ev.push('No default gateway was detected on this interface.');}
if(CAP.hop1===1){sc+=20;ev.push('The first traceroute hop did not answer.');}
if(CAP.gw==='clean'&&CAP.loss>0){ev.push('The gateway itself is clean, so the local router is unlikely to be the origin.');}
s.push({k:'Local Router / Gateway',sc:sc,ev:ev});
sc=0;ev=[];
if(CAP.dnsOk===0){sc+=65;ev.push('DNS resolution for the test domain did not return an answer.');}
if(CAP.httpCode==='000'){sc+=35;ev.push('The HTTPS reachability test did not complete. Name resolution or egress is blocked or timing out.');}
else if(CAP.httpCode!=='200'){sc+=20;ev.push('HTTPS reachability test returned status '+CAP.httpCode+'.');}
if(CAP.dns==='Not detected'){sc+=20;ev.push('No DNS servers were configured on the active interface.');}
s.push({k:'DNS / Name Resolution',sc:sc,ev:ev});
sc=0;ev=[];
if(!ISWIFI&&CAP.loss>0){sc+=65;ev.push('Loss is present on a WIRED connection. This points upstream of the wireless network entirely.');}
if(CAP.gw==='clean'&&CAP.loss>=2){sc+=40;ev.push('The gateway is clean but internet-bound traffic is losing '+CAP.loss+' percent, so the break is beyond the office router.');}
if(CAP.max>250&&CAP.avg<120){sc+=15;ev.push('Latency spiked to '+CAP.max+' ms against a '+CAP.avg+' ms average, which is typical of upstream congestion.');}
if(CAP.avg>90){sc+=15;ev.push('Average latency of '+CAP.avg+' ms is high for a business circuit.');}
s.push({k:'ISP / Upstream Circuit',sc:sc,ev:ev});
sc=0;ev=[];
if(CAP.ssid==='Not available'&&ISWIFI){sc+=10;ev.push('SSID could not be read. This is usually a permissions or driver reporting quirk, not a fault.');}
if(CAP.n<40){sc+=8;ev.push('Only '+CAP.n+' samples were taken. Run an Extended or Deep capture before ruling anything out.');}
if(CAP.loss===0&&CAP.drop===0){ev.push('No endpoint-level packet loss was observed during this capture window.');}
s.push({k:'Endpoint / Client Device',sc:sc,ev:ev});
s.forEach(function(x){if(x.sc>100)x.sc=100;});
s.sort(function(a,b){return b.sc-a.sc;});
return s;}};
var SCORES=Score.run();
var VERDICT=(function(){var top=SCORES[0];
if(CAP.loss===0&&CAP.drop===0&&(!CAP.haveSig||CAP.rssi>-70))return {t:'LINK HEALTHY',p:'NO FAULT OBSERVED',lamp:'lamp-ok',pill:'sp-ok',emo:String.fromCodePoint(128994),cls:'ok'};
if(CAP.loss>=5||CAP.drop>=3)return {t:'LINK DEGRADED',p:top.sc>0?('LIKELY: '+top.k.toUpperCase()):'INVESTIGATE',lamp:'lamp-bad',pill:'sp-bad',emo:String.fromCodePoint(128308),cls:'bad'};
if(CAP.loss>0||(CAP.haveSig&&CAP.rssi<=-72)||CAP.jit>30)return {t:'MARGINAL LINK',p:top.sc>0?('WATCH: '+top.k.toUpperCase()):'MONITOR',lamp:'lamp-warn',pill:'sp-warn',emo:String.fromCodePoint(128993),cls:'warn'};
return {t:'LINK STABLE',p:'WITHIN TOLERANCE',lamp:'lamp-ok',pill:'sp-ok',emo:String.fromCodePoint(128994),cls:'ok'};})();
var Chart={
prep:function(cv,h){var dpr=window.devicePixelRatio||1,w=cv.clientWidth||360;cv.width=w*dpr;cv.height=h*dpr;cv.style.height=h+'px';var c=cv.getContext('2d');c.setTransform(dpr,0,0,dpr,0,0);c.clearRect(0,0,w,h);return {c:c,w:w,h:h};},
css:function(v){return getComputedStyle(document.documentElement).getPropertyValue(v).trim();},
latency:function(cv){var o=this.prep(cv,180),c=o.c,w=o.w,h=o.h;var d=CAP.samples;if(WIN>0&&d.length>WIN)d=d.slice(d.length-WIN);
if(!d.length){c.fillStyle=this.css('--txt-mute');c.font='11px '+this.css('--mono');c.textAlign='center';c.fillText('No samples',w/2,h/2);return;}
var ok=d.filter(function(x){return x.ok;});var mx=ok.length?Math.max.apply(null,ok.map(function(x){return x.ms;})):1;mx=Math.max(mx*1.15,10);
var pad=34,plotH=h-40;
c.strokeStyle='rgba(125,147,168,.13)';c.lineWidth=1;
for(var g=0;g<=4;g++){var y=10+plotH*(g/4);c.beginPath();c.moveTo(pad,y);c.lineTo(w-4,y);c.stroke();c.fillStyle=this.css('--txt-mute');c.font='8px '+this.css('--mono');c.textAlign='right';c.fillText(Math.round(mx*(1-g/4))+'',pad-4,y+3);}
var step=(w-pad-6)/Math.max(1,d.length-1);
d.forEach(function(s,i){if(s.ok)return;var x=pad+i*step;c.fillStyle='rgba(255,59,87,.22)';c.fillRect(x-Math.max(1.5,step/2),10,Math.max(3,step),plotH);});
c.beginPath();var started=false;
d.forEach(function(s,i){var x=pad+i*step;if(!s.ok){started=false;return;}var y=10+plotH-(s.ms/mx)*plotH;if(!started){c.moveTo(x,y);started=true;}else{c.lineTo(x,y);}});
c.strokeStyle=this.css('--cyan');c.lineWidth=1.8;c.stroke();
var avgY=10+plotH-(CAP.avg/mx)*plotH;c.strokeStyle=this.css('--violet');c.setLineDash([4,3]);c.lineWidth=1.1;c.beginPath();c.moveTo(pad,avgY);c.lineTo(w-4,avgY);c.stroke();c.setLineDash([]);
c.fillStyle=this.css('--violet');c.font='700 8px '+this.css('--mono');c.textAlign='left';c.fillText('AVG '+CAP.avg+' ms',pad+3,avgY-4);
c.fillStyle=this.css('--txt-mute');c.font='8px '+this.css('--mono');c.textAlign='left';c.fillText(d[0].t,pad,h-4);c.textAlign='right';c.fillText(d[d.length-1].t,w-4,h-4);c.textAlign='center';c.fillText(d.length+' samples   red bands are failed pings',w/2,h-4);},
signal:function(cv){var o=this.prep(cv,150),c=o.c,w=o.w,h=o.h,self=this;
if(!CAP.haveSig){c.fillStyle=this.css('--txt-mute');c.font='11px '+this.css('--mono');c.textAlign='center';c.fillText('Signal metrics unavailable on this link',w/2,h/2);return;}
var rows=[['RSSI',CAP.rssi,-95,-35,[-72,-67]]];
if(HAVENOISE){rows.push(['NOISE',CAP.noise,-100,-40,[-90,-80]]);rows.push(['SNR',CAP.snr,0,50,[20,25]]);}
if(HASPCT){rows.push(['QUALITY',CAP.sigPct,0,100,[40,60]]);}
var barH=18,gap=28,x0=54,bw=w-x0-64;
rows.forEach(function(r,i){var y=16+i*gap;var t=(r[1]-r[2])/(r[3]-r[2]);if(t<0)t=0;if(t>1)t=1;
c.fillStyle=self.css('--txt-mute');c.font='9px '+self.css('--mono');c.textAlign='left';c.fillText(r[0],4,y+13);
c.fillStyle='rgba(255,255,255,.07)';c.fillRect(x0,y,bw,barH);
var col=self.css('--green');
if(r[0]==='SNR'){if(r[1]<r[4][0])col=self.css('--red');else if(r[1]<r[4][1])col=self.css('--amber');}
else if(r[0]==='RSSI'){if(r[1]<=r[4][0])col=self.css('--red');else if(r[1]<=r[4][1])col=self.css('--amber');}
else if(r[0]==='QUALITY'){if(r[1]<r[4][0])col=self.css('--red');else if(r[1]<r[4][1])col=self.css('--amber');}
else{if(r[1]>r[4][1])col=self.css('--red');else if(r[1]>r[4][0])col=self.css('--amber');}
c.fillStyle=col;c.fillRect(x0,y,bw*t,barH);
c.fillStyle=self.css('--txt');c.font='700 11px '+self.css('--mono');c.textAlign='right';
var unit=' dBm';if(r[0]==='SNR')unit=' dB';if(r[0]==='QUALITY')unit='%';
c.fillText(r[1]+unit,w-4,y+13);});
c.fillStyle=this.css('--txt-mute');c.font='8px '+this.css('--mono');c.textAlign='left';
c.fillText(HAVENOISE?'targets: RSSI better than -67 dBm, noise below -90 dBm, SNR above 25 dB':'targets: RSSI better than -67 dBm, link quality above 60 percent',4,h-6);},
chan:function(cv){var o=this.prep(cv,150),c=o.c,w=o.w,h=o.h,self=this;var d=CAP.nb.slice(0,14);
if(!d.length){c.fillStyle=this.css('--txt-mute');c.font='11px '+this.css('--mono');c.textAlign='center';c.fillText('No neighbouring networks reported',w/2,h/2);return;}
d.sort(function(a,b){return a.ch-b.ch;});
var mx=Math.max.apply(null,d.map(function(x){return x.n;}));var plotH=h-36,bw=(w-14)/d.length;
d.forEach(function(x,i){var bh=(x.n/mx)*plotH,px=7+i*bw,py=12+plotH-bh;var mine=(x.ch===CHAN);
var col=mine?self.css('--cyan'):(x.ch<=14?self.css('--amber'):self.css('--violet'));
var g=c.createLinearGradient(0,py,0,12+plotH);g.addColorStop(0,col);g.addColorStop(1,'rgba(125,147,168,.08)');
c.fillStyle=g;c.fillRect(px+1.5,py,bw-3,Math.max(2,bh));
if(mine){c.strokeStyle='#fff';c.lineWidth=1.4;c.strokeRect(px+1.5,py,bw-3,Math.max(2,bh));}
c.fillStyle=self.css('--txt');c.font='700 9px '+self.css('--mono');c.textAlign='center';c.fillText(x.n+'',px+bw/2,py-4);
c.fillStyle=mine?self.css('--cyan'):self.css('--txt-mute');c.font='8px '+self.css('--mono');c.fillText(x.ch+'',px+bw/2,h-14);});
c.fillStyle=this.css('--txt-mute');c.font='8px '+this.css('--mono');c.textAlign='center';c.fillText('networks per channel   cyan is your channel, amber is 2.4 GHz, violet is 5 or 6 GHz',w/2,h-3);},
compare:function(cv){var o=this.prep(cv,160),c=o.c,w=o.w,h=o.h,self=this;var d=CAPS.slice(-8);
if(d.length<2){c.fillStyle=this.css('--txt-mute');c.font='11px '+this.css('--mono');c.textAlign='center';c.fillText('Run a second capture at another position to compare',w/2,h/2);return;}
var mx=Math.max(1,Math.max.apply(null,d.map(function(x){return x.loss;})));var plotH=h-44,bw=(w-14)/d.length;
d.forEach(function(x,i){var bh=(x.loss/mx)*plotH,px=7+i*bw,py=14+plotH-bh;
var col=x.loss>=5?self.css('--red'):(x.loss>0?self.css('--amber'):self.css('--green'));
c.fillStyle=col;c.fillRect(px+2,py,bw-4,Math.max(2,bh));
if(x.id===CAP.id){c.strokeStyle='#fff';c.lineWidth=1.4;c.strokeRect(px+2,py,bw-4,Math.max(2,bh));}
c.fillStyle=self.css('--txt');c.font='700 9px '+self.css('--mono');c.textAlign='center';c.fillText(fx(x.loss,1)+'%',px+bw/2,py-4);
c.fillStyle=self.css('--txt-mute');c.font='7.5px '+self.css('--mono');
c.fillText(x.loc.split(' ')[0].substring(0,10),px+bw/2,h-14);
c.fillText(x.link.indexOf('Wi-Fi')===0?'wifi':'wired',px+bw/2,h-4);});
c.fillStyle=this.css('--txt-mute');c.font='8px '+this.css('--mono');c.textAlign='left';c.fillText('packet loss by capture point',4,10);}};
var Plain={
grade:function(){
 if(CAP.loss===0&&CAP.drop===0)return 'ok';
 if(CAP.loss>=5||CAP.drop>=3)return 'bad';
 return 'warn';},
headline:function(){
 var g=this.grade();
 if(g==='ok')return {t:'Your connection looked healthy during this test.',s:'We sent '+CAP.n+' small test messages out to the internet and every single one came back. Nothing was lost and nothing was interrupted while we were watching.'};
 if(g==='bad')return {t:'Your connection is genuinely dropping out. This is not your imagination.',s:'We sent '+CAP.n+' small test messages out to the internet. '+CAP.bad+' of them never came back, and at one point '+CAP.drop+' in a row went missing. That gap is exactly the moment a call freezes or a page stops loading.'};
 return {t:'Your connection is mostly working, but it is not solid.',s:'We sent '+CAP.n+' small test messages out to the internet and '+CAP.bad+' did not come back. That is not enough to break everything, but it is enough to cause the occasional stutter on a call or a page that hangs for a second.'};},
plainCause:function(){
 var top=SCORES[0];
 if(top.sc<=0)return 'Nothing in this test points to a specific fault. If it still misbehaves, we need to capture again while the problem is actually happening.';
 var m={
  'RF Interference / Noise':'The airwaves around you are crowded. Your Wi-Fi is competing with a lot of other signals nearby.',
  'Access Point Coverage / Placement':'You are too far from the Wi-Fi equipment, or something solid is in the way.',
  'Roaming / Association Loss':'Your device is briefly letting go of the Wi-Fi and having to reconnect.',
  'Local Router / Gateway':'The problem starts at the office box that connects everyone to the internet.',
  'DNS / Name Resolution':'Your device is struggling to look up website addresses.',
  'ISP / Upstream Circuit':'The office equipment looks fine. The trouble is on the internet line coming into the building.',
  'Endpoint / Client Device':'This looks specific to this one computer rather than the office network.'};
 return m[top.k]||top.k;},
message:function(){
 var g=this.grade(),L=[];
 L.push('Hi,');
 L.push('');
 if(g==='ok'){
  L.push('I ran a network test on the affected machine at the ' + CAP.loc.toLowerCase() + ' and the connection held up cleanly the whole time. Out of ' + CAP.n + ' test messages, none were lost.');
  L.push('');
  L.push('That does not mean the problem is imaginary. It means it was not happening during the window I measured. The next step is to run the same test while the issue is actually occurring so I can catch it in the act.');
 } else if(g==='bad'){
  L.push('I ran a network test on the affected machine at the ' + CAP.loc.toLowerCase() + '. It confirms what you have been experiencing.');
  L.push('');
  L.push('Out of ' + CAP.n + ' test messages sent to the internet, ' + CAP.bad + ' never came back, and at one point ' + CAP.drop + ' in a row were lost. That is the moment a call drops or a page stops loading.');
  L.push('');
  L.push('Most likely cause: ' + this.plainCause());
 } else {
  L.push('I ran a network test on the affected machine at the ' + CAP.loc.toLowerCase() + '. The connection is working, but it is not as steady as it should be.');
  L.push('');
  L.push('Out of ' + CAP.n + ' test messages, ' + CAP.bad + ' did not come back. That is enough to cause an occasional stutter, though not a total outage.');
  L.push('');
  L.push('Most likely cause: ' + this.plainCause());
 }
 L.push('');
 L.push('Next step: ' + (g==='ok'?'re-test during the next reported failure.':'I will repeat this test in a couple of other spots so we can pin down exactly where it breaks down.'));
 L.push('');
 L.push('I will keep you posted.');
 L.push('');
 L.push('Thank you,');
 L.push('Full Name');
 L.push('IT Support');
 return L.join(String.fromCharCode(10));}};
var Panels={hidden:LS.get('hidden',{}),collapsed:LS.get('collapsed',{}),
defs:[{id:'explain',ico:String.fromCodePoint(128172),title:'PLAIN ENGLISH BRIEFING',wide:true,r:'rExplain'},
{id:'verdict',ico:String.fromCodePoint(127919),title:'ROOT CAUSE MATRIX',wide:true,r:'rVerdict'},
{id:'latency',ico:String.fromCodePoint(128200),title:'LATENCY AND DROP TIMELINE',wide:true,r:'rLatency'},
{id:'signal',ico:String.fromCodePoint(128246),title:'SIGNAL QUALITY',wide:false,r:'rSignal'},
{id:'chan',ico:String.fromCodePoint(128225),title:'CHANNEL CONGESTION',wide:false,r:'rChan'},
{id:'compare',ico:String.fromCodePoint(9878),title:'CAPTURE COMPARISON',wide:true,r:'rCompare'},
{id:'path',ico:String.fromCodePoint(128279),title:'PATH VALIDATION',wide:false,r:'rPath'},
{id:'actions',ico:String.fromCodePoint(9989),title:'RECOMMENDED ACTIONS',wide:false,r:'rActions'},
{id:'device',ico:String.fromCodePoint(128187),title:'DEVICE AND INTERFACE',wide:false,r:'rDevice'},
{id:'samples',ico:String.fromCodePoint(128203),title:'SAMPLE LOG',wide:false,r:'rSamples'},
{id:'copilot',ico:String.fromCodePoint(129302),title:'COPILOT PROMPTS',wide:true,r:'rCopilot'},
{id:'raw',ico:String.fromCodePoint(128220),title:'RAW DIAGNOSTIC OUTPUT',wide:true,r:'rRaw'}],
visible:function(){var self=this;return this.defs.filter(function(d){return !self.hidden[d.id];});},
toggle:function(id){this.hidden[id]=!this.hidden[id];LS.set('hidden',this.hidden);this.build();this.chips();},
collapse:function(id){this.collapsed[id]=!this.collapsed[id];LS.set('collapsed',this.collapsed);var el=document.getElementById('panel_'+id);if(el)el.classList.toggle('collapsed',!!this.collapsed[id]);},
focus:function(id){if(this.hidden[id])this.toggle(id);var el=document.getElementById('panel_'+id);if(!el)return;if(this.collapsed[id])this.collapse(id);el.scrollIntoView({behavior:'smooth',block:'center'});el.classList.add('glow');setTimeout(function(){el.classList.remove('glow');},1800);},
chips:function(){var self=this;q('#chipRow').innerHTML=this.defs.map(function(d){return '<button class="chip '+(self.hidden[d.id]?'':'on')+'" data-act="chip" data-id="'+d.id+'">'+d.ico+' '+d.title.split(' ')[0]+'</button>';}).join('');},
build:function(){var self=this;q('#grid').innerHTML=this.visible().map(function(d){return '<section class="panel '+(d.wide?'wide':'')+(self.collapsed[d.id]?' collapsed':'')+'" id="panel_'+d.id+'"><div class="p-head" data-act="collapse" data-id="'+d.id+'"><h3><span class="caret">&#9662;</span><span>'+d.ico+'</span>'+d.title+(d.id==='explain'?' <span class="badge new">FOR EVERYONE</span>':'')+'</h3><button class="p-act" data-act="hide" data-id="'+d.id+'">&#10005;</button></div><div class="p-body" id="body_'+d.id+'"></div></section>';}).join('');this.renderAll();},
render:function(id){var d=null;this.defs.forEach(function(x){if(x.id===id)d=x;});if(!d)return;var b=document.getElementById('body_'+id);if(!b)return;b.innerHTML=Render[d.r]();if(Render[d.r+'_after'])Render[d.r+'_after']();},
renderAll:function(){var self=this;this.visible().forEach(function(d){self.render(d.id);});}};
var Render={
rExplain:function(){
 function blk(cls,title,analogy,result,meaning){return '<div class="xp '+cls+'"><h4>'+esc(title)+'</h4><div class="an">'+esc(analogy)+'</div><div class="rs">YOUR RESULT: '+esc(result)+'</div><div class="mn">'+meaning+'</div></div>';}
 var hd=Plain.headline(),g=Plain.grade();
 var h='<div class="hint" style="margin-bottom:12px">This section explains the whole report in everyday language. No technical background needed. Every number below is from the test that just ran on this machine.</div>';
 h+='<div class="vbox '+g+'"><div class="vt">'+esc(hd.t)+'</div><div class="vs">'+esc(hd.s)+'</div></div>';
 h+='<div class="hint" style="margin:14px 0 9px">WHAT EACH MEASUREMENT MEANS</div>';
 var c1=CAP.loss===0?'ok':(CAP.loss>=5?'bad':'warn');
 h+=blk(c1,'Packet loss - did anything go missing?',
  'Think of posting ' + CAP.n + ' letters and counting how many replies come back. Every letter should get a reply. Any that go missing are gone for good, and the internet simply gives up on them.',
  CAP.bad + ' of ' + CAP.n + ' went missing (' + CAP.loss + '%)',
  CAP.loss===0?'Nothing was lost. This is what a healthy connection looks like.':(CAP.loss>=5?'This is a lot. At this level video calls freeze, audio breaks up and files fail to upload. This is the single clearest sign that something is genuinely wrong.':'A small amount went missing. You would notice this as an occasional hiccup rather than a full outage, but it should be zero.'));
 var c2=CAP.avg<=40?'ok':(CAP.avg>90?'bad':'warn');
 h+=blk(c2,'Latency - how long does a reply take?',
  'This is how quickly someone answers when you call their name across the office. Lower is better. Anything under about 40 milliseconds feels instant to a human being.',
  'average ' + CAP.avg + ' ms (fastest ' + CAP.min + ', slowest ' + CAP.max + ')',
  CAP.avg<=40?'This is quick. Nothing here would make anything feel sluggish.':(CAP.avg>90?'This is slow enough to feel. Pages hesitate before loading and people talk over each other on calls because the delay makes it hard to judge when someone has finished.':'A little slower than ideal, but most people would not consciously notice this on its own.'));
 var c3=CAP.jit<=15?'ok':(CAP.jit>30?'bad':'warn');
 h+=blk(c3,'Jitter - is the rhythm steady?',
  'Imagine a bus that is supposed to come every ten minutes. If it arrives every ten minutes exactly, you can plan around it. If it arrives after two minutes, then eighteen, then five, the average is still ten but it is useless. Jitter measures that unevenness.',
  CAP.jit + ' ms of variation between messages',
  CAP.jit<=15?'The rhythm is steady. Voice and video have a smooth stream of data to work with.':(CAP.jit>30?'This is uneven, and it is the usual reason voices sound robotic or chopped up on Zoom and Teams even when the internet speed test looks fine. Speed is not the problem here; consistency is.':'Slightly uneven. Occasionally noticeable on a call, but not severe.'));
 var c4=CAP.drop===0?'ok':(CAP.drop>=3?'bad':'warn');
 h+=blk(c4,'Longest unbroken gap - did it actually cut out?',
  'There is a difference between a phone call that sounds a bit muffled and one where the line goes completely dead for four seconds. This measures the dead silence, not the muffling.',
  CAP.drop===0?'no gaps at all':CAP.drop + ' messages lost back to back',
  CAP.drop===0?'The connection never actually let go. Whatever else is going on, it did not fully cut out during this test.':(CAP.drop>=3?'This is a real disconnection, not just slowness. This is the moment someone says "I lost you" on a call. It is the most important finding in this report.':'A very brief interruption. Probably not noticeable by itself.'));
 if(ISWIFI&&CAP.haveSig){
  var c5=CAP.rssi>-67?'ok':(CAP.rssi<=-72?'bad':'warn');
  h+=blk(c5,'Signal strength - how clearly can your device hear the Wi-Fi?',
   'Picture someone speaking to you from across a room. Up close they are easy to hear. From the far end, through a closed door, you catch maybe half the words and have to keep asking them to repeat themselves. Your laptop has the same problem with the Wi-Fi box.',
   CAP.rssi + ' dBm' + (HASPCT?(' (about ' + CAP.sigPct + '% quality)'):'') + ' on channel ' + esc(CAP.chan) + ', ' + BAND,
   CAP.rssi>-67?'Strong. Your device is hearing the Wi-Fi clearly from where it is sitting.':(CAP.rssi<=-72?'Weak. Your device is straining to hear the Wi-Fi from this spot. The simplest test is to sit closer to the Wi-Fi box and see if things improve immediately.':'Usable but not strong. Moving closer would help.'));
 }
 if(ISWIFI&&HAVENOISE){
  var c6=CAP.snr>=25?'ok':'warn';
  h+=blk(c6,'Background noise - how noisy is the room?',
   'You can hear someone perfectly in a quiet library. Say the exact same words at the same volume in a busy restaurant and you catch nothing. The voice did not get quieter; the background got louder. This measures the background.',
   'signal is ' + CAP.snr + ' dB above the background noise',
   CAP.snr>=25?'Comfortably above the noise. Your device can pick out the Wi-Fi easily.':'The background is loud relative to your signal. This is the classic fingerprint of a shared office building, where all the neighbouring businesses are competing for the same airwaves.');
 }
 if(ISWIFI){
  var c7=CAP.nbTotal>12?'bad':(CAP.nbTotal>6?'warn':'ok');
  h+=blk(c7,'Neighbouring networks - how many others are competing?',
   'Wi-Fi channels are like lanes on a road. If you have a lane to yourself you move freely. If forty cars are merging into your lane from the businesses on either side, everyone crawls, no matter how good your car is.',
   CAP.nbTotal + ' other networks detected nearby, you are on channel ' + CHAN + ' (' + BAND + ')',
   CAP.nbTotal>12?'That is a crowded neighbourhood. In a shared building this is one of the most common reasons Wi-Fi is fine one minute and unusable the next. The fix is usually changing which lane the office Wi-Fi uses, which the landlord or network vendor has to do.':(CAP.nbTotal>6?'A moderate amount of competition. Worth noting, not alarming on its own.':'Not many competitors nearby. Congestion is unlikely to be your problem.'));
 }
 var c8=CAP.gw==='clean'?'ok':(CAP.gw==='unknown'?'warn':'bad');
 h+=blk(c8,'The office box - is the problem inside or outside the building?',
  'Every office has one box that everything goes through on its way to the internet, like the single front door of a building. We knocked on that door separately. If the door answers fine but your post still goes missing, the problem is out on the road, not in the building.',
  CAP.gw==='clean'?'the office box answered every time':(CAP.gw==='lossy'?'the office box answered, but not reliably':'the office box did not answer properly'),
  CAP.gw==='clean'?(CAP.loss>0?'The office box itself is healthy, so the trouble is further out - either the Wi-Fi in the air between you and it, or the internet line leaving the building.':'The office box is healthy and nothing was lost. Everything checks out here.'):'The office box is not responding cleanly, which means the problem starts inside the building. That points at the router or the local network equipment rather than the internet provider.');
 var c9=CAP.dnsOk===1?'ok':'bad';
 h+=blk(c9,'Address lookup - can your device find websites by name?',
  'You know the name of the person you want to call, but you still need the phone book to get their number. Your device does the same thing every time you type a website name. If the phone book is missing, everything looks broken even when the connection itself is perfect.',
  CAP.dnsOk===1?'lookups are working normally':'lookups did not come back',
  CAP.dnsOk===1?'Your device can find websites by name without any trouble.':'This alone can make the internet appear completely dead even when the connection is fine. It is often a quick fix and is worth checking before anything else.');
 h+='<div class="hint" style="margin:16px 0 9px">WHAT HAPPENS NEXT</div>';
 var nx=[];
 if(ISWIFI&&CAP.loss>0)nx.push('Run this same test on a cable instead of Wi-Fi from the same desk. That one comparison tells us whether to chase the Wi-Fi or the internet line, and it saves days of guessing.');
 if(ISWIFI&&CAP.haveSig&&CAP.rssi<=-72)nx.push('Run the test again sitting right next to the Wi-Fi box. If the numbers jump, it is a coverage problem and the answer is another access point, not a new laptop.');
 if(!ISWIFI&&CAP.loss>0)nx.push('This test ran on a cable and still lost data, which means the Wi-Fi is not to blame. This becomes a conversation with the internet provider, and this report is the evidence.');
 if(CAP.nbTotal>12)nx.push('Share the neighbouring network count with the building or the network vendor. It is the justification for changing the Wi-Fi channel plan.');
 if(CAP.loss===0&&CAP.drop===0)nx.push('Run this again the moment the problem happens next. A clean test only proves it was behaving while we watched, and catching it misbehaving is what we need.');
 nx.push('Keep this report. Every test is saved and compared automatically, so the picture gets clearer with each one rather than starting from scratch.');
 h+='<ul style="margin:0;padding-left:18px">'+nx.map(function(x){return '<li style="margin:7px 0;line-height:1.6">'+esc(x)+'</li>';}).join('')+'</ul>';
 h+='<div class="hint" style="margin:16px 0 9px">READY-TO-SEND UPDATE</div>';
 h+='<div class="hint" style="margin-bottom:8px">Written for a non-technical reader. Copy it straight into Zoom, Teams or email.</div>';
 h+='<div class="row noprint" style="margin-bottom:8px"><button class="btn" data-act="copyMsg">COPY THIS MESSAGE</button></div>';
 h+='<div class="msgbox">'+esc(Plain.message())+'</div>';
 h+='<div class="hint" style="margin:16px 0 9px">JARGON DECODER</div>';
 h+='<div class="glos">'+
  '<b>Packet</b><span>One small piece of internet data. Everything you do online is chopped into thousands of these.</span>'+
  '<b>Packet loss</b><span>Pieces that never arrived. The percentage of things that went missing.</span>'+
  '<b>Latency / ms</b><span>Delay, measured in milliseconds. One thousand milliseconds is one second.</span>'+
  '<b>Jitter</b><span>How uneven the delay is. Steady is good, erratic is bad, even at the same average speed.</span>'+
  '<b>RSSI / dBm</b><span>How loudly your device hears the Wi-Fi. Closer to zero is stronger, so -55 is much better than -85.</span>'+
  (HAVENOISE?'<b>SNR</b><span>How far your signal stands above the background noise. Higher is clearer.</span>':'')+
  '<b>Channel</b><span>Which lane the Wi-Fi is using. Neighbours on the same lane slow each other down.</span>'+
  '<b>Gateway</b><span>The office box everything passes through on the way to the internet.</span>'+
  '<b>DNS</b><span>The internet phone book that turns a website name into an address.</span>'+
  '<b>Access point</b><span>The Wi-Fi transmitter on the wall or ceiling that your device connects to.</span>'+
  '</div>';
 return h;},
rVerdict:function(){var h='<div class="hint" style="margin-bottom:10px">Weighted ranking of where to investigate first, computed from this capture only. Bars show relative confidence, not probability. For a non-technical explanation, see the Plain English Briefing panel.</div>';
SCORES.forEach(function(s){var col=s.sc>=50?'--red':(s.sc>=25?'--amber':(s.sc>0?'--violet':'--green'));
h+='<div class="rc"><div class="rc-top"><b>'+esc(s.k)+'</b><span class="mono" style="color:var('+col+')">'+s.sc+'</span></div><div class="bar"><i style="width:'+Math.max(2,s.sc)+'%;background:var('+col+')"></i></div><div class="ev">'+(s.ev.length?s.ev.map(function(e){return '&bull; '+esc(e);}).join('<br>'):'&bull; No supporting evidence in this capture.')+'</div></div>';});
var top=SCORES[0];var cls=top.sc>=50?'bad':(top.sc>=25?'warn':'good');
var msg=top.sc<=0?'Nothing in this capture indicates a fault. If the user still reports drops, repeat with a Deep capture during an active Zoom or Teams call.':('Start with <b>'+esc(top.k)+'</b>. '+esc(top.ev[0]||''));
return h+'<div class="advice '+cls+'">'+msg+'</div>';},
rLatency:function(){return '<canvas id="cvLat"></canvas><div class="result"><div class="line"><span>Samples sent</span><b>'+CAP.n+' to '+esc(CAP.target)+'</b></div><div class="line"><span>Replies / timeouts</span><b><span class="c-green">'+CAP.ok+'</span> / <span class="c-red">'+CAP.bad+'</span></b></div><div class="line"><span>Packet loss</span><b class="'+(CAP.loss>0?'c-red':'c-green')+'">'+CAP.loss+'%</b></div><div class="line"><span>Latency min / avg / max</span><b>'+CAP.min+' / '+CAP.avg+' / '+CAP.max+' ms</b></div><div class="line"><span>Jitter (mean deviation)</span><b class="'+(CAP.jit>30?'c-amber':'')+'">'+CAP.jit+' ms</b></div><div class="line"><span>Longest unbroken outage</span><b class="'+(CAP.drop>=3?'c-red':'')+'">'+CAP.drop+' samples</b></div></div><div class="hint" style="margin-top:8px">A single failed sample is usually noise. Three or more in a row is a real link drop, and that is what users describe as the connection cutting out.</div>';},
rLatency_after:function(){var cv=document.getElementById('cvLat');if(cv)Chart.latency(cv);},
rSignal:function(){var h='<canvas id="cvSig"></canvas><div class="result"><div class="line"><span>SSID</span><b>'+esc(CAP.ssid)+'</b></div><div class="line"><span>BSSID (radio)</span><b>'+esc(CAP.bssid)+'</b></div><div class="line"><span>Channel / band</span><b>'+esc(CAP.chan)+' &middot; '+BAND+'</b></div><div class="line"><span>PHY mode</span><b>'+esc(CAP.phy)+'</b></div><div class="line"><span>Security</span><b>'+esc(CAP.sec)+'</b></div><div class="line"><span>Transmit rate</span><b>'+esc(CAP.tx)+'</b></div>'+(HASPCT?'<div class="line"><span>Link quality</span><b>'+CAP.sigPct+'%</b></div>':'')+(HAVENOISE?'':'<div class="line"><span>Noise / SNR</span><b class="mute">not reported on '+esc(PLAT)+'</b></div>')+'</div>';
if(!ISWIFI)h+='<div class="advice good">This capture ran over <b>'+esc(CAP.link)+'</b>. Comparing it against a Wi-Fi capture from the same desk is the fastest way to separate wireless problems from router or ISP problems.</div>';
else if(CAP.haveSig&&CAP.rssi<=-72)h+='<div class="advice bad">Signal is weak at this position. Repeat the capture standing next to the access point. If the numbers improve sharply this is a coverage or placement problem, not a device problem.</div>';
else if(HAVENOISE&&CAP.snr<25)h+='<div class="advice warn">Signal strength is acceptable but the noise floor is high, which is the classic signature of a shared multi-tenant building.</div>';
else if(!HAVENOISE&&ISWIFI)h+='<div class="advice warn">Signal strength is acceptable at this position. '+esc(PLAT)+' does not expose a noise floor, so run the macOS build on a Mac at the same desk if you need SNR to confirm or rule out interference.</div>';
else h+='<div class="advice good">Signal and noise are within healthy limits at this position.</div>';
return h;},
rSignal_after:function(){var cv=document.getElementById('cvSig');if(cv)Chart.signal(cv);},
rChan:function(){var h='<canvas id="cvChan"></canvas>';var same=0;CAP.nb.forEach(function(x){if(x.ch===CHAN)same=x.n;});
h+='<div class="result"><div class="line"><span>Neighbouring networks</span><b class="'+(CAP.nbTotal>12?'c-red':(CAP.nbTotal>6?'c-amber':'c-green'))+'">'+CAP.nbTotal+'</b></div><div class="line"><span>Sharing your channel ('+CHAN+')</span><b class="'+(same>2?'c-red':'')+'">'+same+'</b></div><div class="line"><span>Your band</span><b>'+BAND+'</b></div></div>';
if(CAP.nbTotal>12)h+='<div class="advice bad">Dense RF environment. In a shared building this is a leading cause of intermittent Wi-Fi. Move the office SSID to a clear 5 GHz or 6 GHz channel and disable low legacy data rates.</div>';
else if(BAND==='2.4 GHz')h+='<div class="advice warn">The client is associated on 2.4 GHz. Steer these clients to 5 GHz. 2.4 GHz has only three non-overlapping channels and is shared with every neighbour, microwave and Bluetooth device nearby.</div>';
else h+='<div class="advice good">Channel occupancy at this position looks reasonable.</div>';
return h;},
rChan_after:function(){var cv=document.getElementById('cvChan');if(cv)Chart.chan(cv);},
rCompare:function(){var h='<canvas id="cvCmp"></canvas>';
if(CAPS.length<2)return h+'<div class="empty">Only this capture is stored.<br>Run the tool again at another position - desk, next to the access point, then on wired Ethernet - and they will line up here automatically.</div>';
h+='<table><thead><tr><th>Capture point</th><th>Link</th><th style="text-align:right">Loss</th><th style="text-align:right">Avg</th><th style="text-align:right">Jitter</th><th style="text-align:right">RSSI</th><th style="text-align:right">SNR</th><th style="text-align:right">Drop</th></tr></thead><tbody>';
CAPS.slice().reverse().forEach(function(x){var cls=x.id===CAP.id?'best':(x.loss>=5?'bad':'');var d=new Date(x.ts);
h+='<tr class="'+cls+'"><td><b>'+esc(x.loc)+'</b>'+(x.id===CAP.id?' <span class="badge new">THIS</span>':'')+'<br><span class="mute" style="font-size:9px">'+d.toLocaleDateString()+' '+d.toLocaleTimeString([],{hour:'2-digit',minute:'2-digit'})+(x.plat?' &middot; '+esc(x.plat):'')+'</span></td><td class="dim">'+esc(x.link)+'</td><td style="text-align:right" class="'+(x.loss>0?'c-red':'c-green')+'"><b>'+x.loss+'%</b></td><td style="text-align:right">'+x.avg+' ms</td><td style="text-align:right">'+x.jit+' ms</td><td style="text-align:right">'+(x.haveSig?x.rssi:'--')+'</td><td style="text-align:right">'+((x.haveNoise===undefined?x.haveSig:x.haveNoise)?x.snr:'--')+'</td><td style="text-align:right" class="'+(x.drop>=3?'c-red':'')+'">'+x.drop+'</td></tr>';});
h+='</tbody></table>';
var wired=CAPS.filter(function(x){return x.link.indexOf('Wi-Fi')!==0;});
var wifi=CAPS.filter(function(x){return x.link.indexOf('Wi-Fi')===0;});
if(wired.length&&wifi.length){var ww=Math.max.apply(null,wired.map(function(x){return x.loss;}));var fw=Math.max.apply(null,wifi.map(function(x){return x.loss;}));
if(ww===0&&fw>0)h+='<div class="advice bad"><b>Wired is clean, Wi-Fi is not.</b> The fault is wireless or environmental - access point coverage, RF interference or channel congestion. The circuit and the endpoint devices are exonerated.</div>';
else if(ww>0)h+='<div class="advice bad"><b>Loss is present on wired as well.</b> This is upstream of the wireless network. Escalate to the ISP and the office router with these captures attached.</div>';
else h+='<div class="advice good">Both wired and wireless captures are clean. Repeat during the reported failure window with a Deep capture.</div>';}
else h+='<div class="advice warn">You have not captured both a wireless and a wired sample yet. That single comparison is what separates a Wi-Fi problem from an ISP problem. Plug in an Ethernet adapter and run the Wired Ethernet Test.</div>';
return h+'<button class="btn ghost sm noprint" style="margin-top:9px" data-act="clearLedger">CLEAR CAPTURE LEDGER</button>';},
rCompare_after:function(){var cv=document.getElementById('cvCmp');if(cv)Chart.compare(cv);},
rPath:function(){var rows=[['Default gateway',CAP.router,CAP.gw==='clean'?'ok':(CAP.gw==='lossy'?'warn':'bad'),CAP.gw==='clean'?'0% loss':(CAP.gw==='lossy'?'responding with loss':'no response')],['DNS resolution',CAP.dns,CAP.dnsOk===1?'ok':'bad',CAP.dnsOk===1?'answered':'no answer'],['HTTPS reachability','connectivity test endpoint',CAP.httpCode==='200'?'ok':'bad','HTTP '+CAP.httpCode],['First traceroute hop',CAP.router,CAP.hop1===0?'ok':'warn',CAP.hop1===0?'answered':'silent'],['Internet echo',CAP.target,CAP.loss===0?'ok':(CAP.loss<5?'warn':'bad'),CAP.loss+'% loss']];
var h='<table><thead><tr><th>Check</th><th>Target</th><th>Result</th></tr></thead><tbody>';
rows.forEach(function(r){var pill=r[2]==='ok'?'<span class="badge ok">PASS</span>':(r[2]==='warn'?'<span class="badge">WATCH</span>':'<span class="badge bad">FAIL</span>');
h+='<tr class="'+(r[2]==='bad'?'bad':'')+'"><td><b>'+esc(r[0])+'</b></td><td class="dim">'+esc(r[1])+'</td><td>'+pill+' <span class="mute">'+esc(r[3])+'</span></td></tr>';});
return h+'</tbody></table><div class="hint" style="margin-top:9px">Reading order: if the gateway passes but the internet echo fails, the break is beyond the office router. If the gateway itself fails, stay inside the office.</div>';},
rActions:function(){var a=[];
if(!ISWIFI&&CAP.loss>0)a.push('Loss on a wired link. Collect these results and open a circuit ticket with the ISP, including the traceroute output below.');
if(ISWIFI&&CAP.loss>0)a.push('Repeat this capture on wired Ethernet from the same desk. That single test decides whether this is a wireless problem or a circuit problem.');
if(ISWIFI&&CAP.haveSig&&CAP.rssi<=-72)a.push('Repeat the capture standing next to the access point and compare the two entries in the comparison panel.');
if(CAP.nbTotal>12)a.push('Document the neighbouring network count for the landlord or network vendor. It is the evidence for a channel plan change.');
if(BAND==='2.4 GHz')a.push('Move the affected clients onto the 5 GHz SSID, or enable band steering if the access point supports it.');
if(CAP.drop>=3)a.push('Check access point roaming and coverage overlap. The client is losing association, not just slowing down.');
if(CAP.dnsOk===0)a.push('Validate the DNS servers on the active interface before assuming a connectivity fault.');
if(CAP.n<60)a.push('Run an Extended or Deep capture during the reported failure window before drawing a conclusion.');
a.push('Capture the same three points on every affected machine: desk, near access point, wired. Consistent results across users point at the environment, not the endpoints.');
a.push('Send the Plain English Briefing to the requester so they understand the finding without needing a translation.');
a.push('Attach the exported JSON and this HTML report to the ServiceNow incident as evidence.');
return '<ul style="margin:0;padding-left:18px">'+a.map(function(x){return '<li style="margin:7px 0;line-height:1.55">'+esc(x)+'</li>';}).join('')+'</ul>';},
rDevice:function(){return '<div class="result" style="margin-top:0"><div class="line"><span>Computer</span><b>'+esc(CAP.host)+'</b></div><div class="line"><span>Console user</span><b>'+esc(CAP.user)+'</b></div><div class="line"><span>Model</span><b>'+esc(CAP.model)+' ('+esc(CAP.modelId)+')</b></div><div class="line"><span>Chip</span><b>'+esc(CAP.chip)+'</b></div><div class="line"><span>Serial</span><b>'+esc(CAP.serial)+'</b></div><div class="line"><span>Operating system</span><b>'+esc(CAP.osName)+' '+esc(CAP.osVer)+' ('+esc(CAP.osBuild)+')</b></div><div class="line"><span>Platform</span><b>'+esc(PLAT)+'</b></div><div class="line"><span>Active interface</span><b>'+esc(CAP.iface)+' &middot; '+esc(CAP.link)+'</b></div><div class="line"><span>Wi-Fi interface</span><b>'+esc(CAP.wifiDev)+'</b></div><div class="line"><span>IP address</span><b>'+esc(CAP.ip)+'</b></div><div class="line"><span>Default gateway</span><b>'+esc(CAP.router)+'</b></div><div class="line"><span>DNS servers</span><b>'+esc(CAP.dns)+'</b></div><div class="line"><span>Capture point</span><b class="c-cyan">'+esc(CAP.loc)+'</b></div><div class="line"><span>Output folder</span><b style="font-size:10px">'+esc(CAP.out)+'</b></div></div>';},
rSamples:function(){var d=CAP.samples.slice();if(WIN>0&&d.length>WIN)d=d.slice(d.length-WIN);d=d.reverse();
var h='<table><thead><tr><th>#</th><th>Time</th><th>Status</th><th style="text-align:right">Latency</th></tr></thead><tbody>';
d.slice(0,120).forEach(function(s){h+='<tr class="'+(s.ok?'':'bad')+'"><td class="mute">'+s.i+'</td><td class="dim">'+esc(s.t)+'</td><td>'+(s.ok?'<span class="c-green">reply</span>':'<span class="c-red">timeout</span>')+'</td><td style="text-align:right"><b>'+(s.ok?s.ms+' ms':'--')+'</b></td></tr>';});
return h+'</tbody></table><div class="hint" style="margin-top:8px">Newest first. The complete sample set is in PingSamples.csv in the output folder.</div>';},
rCopilot:function(){
var sigTxt=HAVENOISE?(', noise '+CAP.noise+' dBm, SNR '+CAP.snr+' dB'):(HASPCT?(', link quality '+CAP.sigPct+' percent, noise floor and SNR are not reported on '+PLAT):(', noise floor and SNR are not reported on '+PLAT));
var p1='Act as a senior enterprise network engineer supporting executives. Analyse this '+PLAT+' Wi-Fi capture and rank the root cause. Capture point: '+CAP.loc+'. Link: '+CAP.link+'. Device: '+CAP.model+' '+CAP.chip+', '+CAP.osName+' '+CAP.osVer+'. SSID '+CAP.ssid+', channel '+CAP.chan+' ('+BAND+'), PHY '+CAP.phy+', transmit rate '+CAP.tx+'. RSSI '+CAP.rssi+' dBm'+sigTxt+'. Neighbouring networks '+CAP.nbTotal+'. Ping to '+CAP.target+': '+CAP.n+' samples, '+CAP.ok+' replies, '+CAP.bad+' timeouts, '+CAP.loss+' percent loss, latency min avg max '+CAP.min+' '+CAP.avg+' '+CAP.max+' ms, jitter '+CAP.jit+' ms, longest unbroken outage '+CAP.drop+' samples. Gateway '+CAP.router+' status '+CAP.gw+'. DNS answered: '+(CAP.dnsOk===1?'yes':'no')+'. HTTPS status '+CAP.httpCode+'. Give me: 1) a ranked root cause analysis with the evidence for each, covering RF interference, access point coverage, channel congestion, roaming, local router, DNS, ISP and endpoint; 2) the exact next tests to run onsite; 3) what to tell the landlord or network vendor if this turns out to be environmental.';
var p2='Using the capture data below, write a ServiceNow incident update for Executive Support. Include an Executive Impact and Business Value statement, a technical Work Notes section listing the investigation performed and the measurements taken, a Resolution or Action Plan with clear next steps, and a Status line. Then write a short, warm, non-technical message for the executive assistant explaining what was found and what happens next, with no IT jargon. Capture: point '+CAP.loc+', link '+CAP.link+', platform '+PLAT+', packet loss '+CAP.loss+' percent, average latency '+CAP.avg+' ms, jitter '+CAP.jit+' ms, longest outage '+CAP.drop+' samples, RSSI '+CAP.rssi+' dBm'+(HAVENOISE?(', SNR '+CAP.snr+' dB'):'')+', '+CAP.nbTotal+' neighbouring networks, channel '+CAP.chan+', gateway status '+CAP.gw+', top ranked cause '+SCORES[0].k+'.';
var p3='Explain this network test result to someone with no technical background at all. Use plain everyday language and a simple analogy for each measurement, the way you would explain it to a busy executive who just wants to know whether it is fixed. Do not use jargon, and if you must use a term, explain it in the same sentence. Keep it under 250 words, warm and reassuring in tone, and end with what happens next. The results: out of '+CAP.n+' test messages sent to the internet, '+CAP.bad+' were lost ('+CAP.loss+' percent). Average response time '+CAP.avg+' ms, with the slowest at '+CAP.max+' ms. Unevenness between messages (jitter) '+CAP.jit+' ms. The longest run of consecutive lost messages was '+CAP.drop+'. '+(ISWIFI?('Wi-Fi signal strength '+CAP.rssi+' dBm on channel '+CAP.chan+' ('+BAND+'), with '+CAP.nbTotal+' other networks detected nearby. '):'This test ran over a wired cable rather than Wi-Fi. ')+'The office gateway status was '+CAP.gw+' and address lookups '+(CAP.dnsOk===1?'worked':'failed')+'. The most likely cause identified was '+SCORES[0].k+'. Location tested: '+CAP.loc+'.';
window.__P1=p1;window.__P2=p2;window.__P3=p3;
return '<div class="hint" style="margin-bottom:9px">Pre-filled with the measurements from this capture. Paste straight into Copilot.</div><div class="row noprint" style="margin-bottom:9px"><button class="btn" data-act="copy1">COPY ROOT CAUSE PROMPT</button><button class="btn" data-act="copy2">COPY SERVICENOW PROMPT</button><button class="btn" data-act="copy3">COPY PLAIN ENGLISH PROMPT</button></div><div class="hint" style="margin-bottom:5px">1. ROOT CAUSE ANALYSIS</div><pre>'+esc(p1)+'</pre><div class="hint" style="margin:11px 0 5px">2. SERVICENOW WORK NOTES</div><pre>'+esc(p2)+'</pre><div class="hint" style="margin:11px 0 5px">3. PLAIN ENGLISH SUMMARY FOR THE END USER</div><pre>'+esc(p3)+'</pre>';},
rRaw:function(){var b=CAP.raw;var secs=[['Gateway test',b.gw],['DNS resolution',b.dns],['HTTPS timing',b.http],['Traceroute',b.trace],['Routing table',b.route],['Interface detail',b.ifc],['Host network configuration',b.scutil],['Preferred wireless networks',b.pref],['Adapter and driver detail',b.air]];
return secs.map(function(s){return '<div class="hint" style="margin:10px 0 4px">'+esc(s[0]).toUpperCase()+'</div><pre>'+esc(s[1]||'(no output)')+'</pre>';}).join('');}};
var Palette={items:[],
open:function(){this.items=[];var self=this;
Panels.defs.forEach(function(d){self.items.push({t:'Go to '+d.title,s:'PANEL',f:function(){Panels.focus(d.id);}});});
this.items.push({t:'Copy the ready-to-send status message',s:'M',f:function(){App.copy(Plain.message(),'Status message copied');}});
this.items.push({t:'Toggle dark / light theme',s:'T',f:function(){App.theme();}});
this.items.push({t:'Export capture snapshot (JSON)',s:'E',f:function(){App.exportJson();}});
this.items.push({t:'Copy root cause prompt',s:'C',f:function(){App.copy(window.__P1,'Root cause prompt copied');}});
this.items.push({t:'Copy ServiceNow prompt',s:'',f:function(){App.copy(window.__P2,'ServiceNow prompt copied');}});
this.items.push({t:'Copy plain English prompt',s:'',f:function(){App.copy(window.__P3,'Plain English prompt copied');}});
this.items.push({t:'Print / save as PDF',s:'P',f:function(){window.print();}});
this.items.push({t:'Clear capture ledger',s:'',f:function(){Ledger.clear();setTimeout(function(){location.reload();},400);}});
this.items.push({t:'Keyboard shortcuts',s:'?',f:function(){App.help();}});
q('#palette').classList.add('show');q('#cmdInput').value='';this.list('');setTimeout(function(){q('#cmdInput').focus();},40);},
close:function(){q('#palette').classList.remove('show');},
list:function(f){f=(f||'').toLowerCase();var self=this;var r=this.items.filter(function(x){return x.t.toLowerCase().indexOf(f)>=0;});
q('#cmdList').innerHTML=r.map(function(x,i){return '<div class="cmd-item'+(i===0?' sel':'')+'" data-act="cmd" data-i="'+self.items.indexOf(x)+'">'+esc(x.t)+'<span class="cs">'+esc(x.s)+'</span></div>';}).join('')||'<div class="empty">No matches</div>';},
runSel:function(){var el=q('#cmdList .cmd-item.sel');if(!el)return;var i=parseInt(el.getAttribute('data-i'),10);this.close();this.items[i].f();},
move:function(dir){var list=Array.prototype.slice.call(document.querySelectorAll('#cmdList .cmd-item'));if(!list.length)return;var i=0;list.forEach(function(x,k){if(x.classList.contains('sel'))i=k;});list[i].classList.remove('sel');var n=i+dir;if(n<0)n=list.length-1;if(n>=list.length)n=0;list[n].classList.add('sel');list[n].scrollIntoView({block:'nearest'});}};
var App={
theme:function(){var h=document.documentElement;var cur=h.getAttribute('data-theme')==='light'?'dark':'light';h.setAttribute('data-theme',cur);LS.set('theme',cur);Panels.renderAll();},
help:function(){q('#helpModal').classList.add('show');},
closeHelp:function(){q('#helpModal').classList.remove('show');},
full:function(){if(document.fullscreenElement)document.exitFullscreen();else document.documentElement.requestFullscreen();},
copy:function(t,m){try{navigator.clipboard.writeText(t);toast(m||'Copied','ok');}catch(e){toast('Copy was blocked by the browser','err');}},
exportJson:function(){var blob=new Blob([JSON.stringify({capture:CAP,scores:SCORES,ledger:CAPS,plainSummary:Plain.message()},null,2)],{type:'application/json'});var a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='WiFiCapture_'+CAP.id+'.json';a.click();toast('Snapshot exported','ok');},
setWin:function(v){WIN=v;LS.set('win',v);Array.prototype.slice.call(document.querySelectorAll('#rangePills .pill')).forEach(function(p){p.classList.toggle('on',parseInt(p.getAttribute('data-win'),10)===v);});Panels.render('latency');Panels.render('samples');},
hero:function(){
q('#lamp').className='statuslamp '+VERDICT.lamp;q('#lamp').textContent=VERDICT.emo;
q('#heroTitle').textContent=VERDICT.t;
q('#heroSub').textContent=CAP.ssid+'   '+CAP.chan+'   '+BAND+'   '+CAP.link+'   '+CAP.n+' samples';
q('#heroBig').textContent=CAP.loss+'%';
q('#heroBig').style.color=CAP.loss>=5?'var(--red)':(CAP.loss>0?'var(--amber)':'var(--green)');
var hp=q('#heroPill');hp.className='statuspill '+VERDICT.pill;hp.textContent=VERDICT.p;
var lp=q('#locPill');lp.textContent=CAP.loc.toUpperCase();
var kp=q('#linkPill');kp.textContent=CAP.link.toUpperCase();kp.className='statuspill '+(ISWIFI?'sp-warn':'sp-ok');
var pp=q('#platPill');if(pp)pp.textContent=PLAT.toUpperCase();
var pn=q('#platNote');if(pn)pn.innerHTML=HAVENOISE?'Signal, channel and neighbour data are read from system_profiler SPAirPortDataType. If SSID reads as unavailable, grant this app Location Services access - all other metrics are unaffected.':'Signal, channel and neighbour data are read from netsh wlan. Windows reports signal as a link-quality percentage; RSSI is derived from it using the standard quality-to-dBm conversion. Windows does not expose a noise floor, so SNR is not shown - run the macOS capture on a Mac at the same desk when SNR is needed.';
var d=new Date(CAP.ts);
q('#clockLocal').textContent=d.toLocaleTimeString([],{hour:'2-digit',minute:'2-digit',second:'2-digit'});
q('#clockMeta').textContent=d.toLocaleDateString([],{weekday:'short',month:'short',day:'numeric'})+'  CAPTURED';
var kpi=[['PACKET LOSS',CAP.loss+'%',CAP.bad+' of '+CAP.n+' failed',CAP.loss>0?'--red':'--green'],
['AVG LATENCY',CAP.avg+' ms','min '+CAP.min+' / max '+CAP.max,CAP.avg>90?'--amber':'--cyan'],
['JITTER',CAP.jit+' ms','mean deviation',CAP.jit>30?'--amber':'--cyan'],
['LONGEST DROP',CAP.drop+'','consecutive timeouts',CAP.drop>=3?'--red':'--green'],
['RSSI',CAP.haveSig?CAP.rssi+' dBm':'--',CAP.haveSig?(CAP.rssi>-67?'strong':(CAP.rssi>-72?'acceptable':'weak')):'not available',CAP.haveSig?(CAP.rssi>-72?'--green':'--red'):'--txt-mute'],
HAVENOISE?['SNR',CAP.snr+' dB',CAP.snr>=25?'clean':'noisy',CAP.snr>=25?'--green':'--amber']:['LINK QUALITY',HASPCT?CAP.sigPct+'%':'--',HASPCT?(CAP.sigPct>=60?'good':'weak'):'SNR not reported',HASPCT?(CAP.sigPct>=60?'--green':'--amber'):'--txt-mute'],
['NEARBY NETWORKS',CAP.nbTotal+'','RF density',CAP.nbTotal>12?'--red':(CAP.nbTotal>6?'--amber':'--green')],
['GATEWAY',CAP.gw.toUpperCase(),CAP.router,CAP.gw==='clean'?'--green':'--red']];
q('#kpis').innerHTML=kpi.map(function(k){return '<div class="kpi"><div class="k">'+k[0]+'</div><div class="v" style="color:var('+k[3]+')">'+esc(k[1])+'</div><div class="d">'+esc(k[2])+'</div></div>';}).join('');
var ab=q('#alertBanner');
if(CAP.loss>=5||CAP.drop>=3){ab.classList.add('show');q('#alertText').textContent='LINK DEGRADED - '+CAP.loss+'% packet loss, longest outage '+CAP.drop+' samples. Top ranked cause: '+SCORES[0].k+'.';}
else if(CAP.loss>0||(CAP.haveSig&&CAP.rssi<=-72)){ab.classList.add('show');q('#alertText').textContent='MARGINAL LINK - review the root cause matrix and repeat this capture at a second position.';}
else ab.classList.remove('show');},
init:function(){var th=LS.get('theme','dark');if(th==='light')document.documentElement.setAttribute('data-theme','light');
this.hero();Panels.chips();Panels.build();this.setWin(WIN);
toast('Capture stored. '+CAPS.length+' capture'+(CAPS.length===1?'':'s')+' in the comparison ledger.','ok',5200);
setTimeout(function(){toast('New: open the Plain English Briefing panel to share these results with a non-technical colleague.','ok',7000);},1400);}};
document.addEventListener('click',function(e){
if(e.target.id==='palette'){Palette.close();return;}
if(e.target.id==='helpModal'){App.closeHelp();return;}
var t=e.target.closest('[data-act]');if(!t)return;
var a=t.getAttribute('data-act'),id=t.getAttribute('data-id');
if(a==='palette')Palette.open();
else if(a==='explain')Panels.focus('explain');
else if(a==='theme')App.theme();
else if(a==='rerender'){Panels.renderAll();toast('Panels re-rendered','ok',2200);}
else if(a==='export')App.exportJson();
else if(a==='print')window.print();
else if(a==='full')App.full();
else if(a==='dismiss')q('#alertBanner').classList.remove('show');
else if(a==='chip')Panels.toggle(id);
else if(a==='collapse')Panels.collapse(id);
else if(a==='hide'){e.stopPropagation();Panels.toggle(id);}
else if(a==='closeHelp')App.closeHelp();
else if(a==='cmd')Palette.runSel();
else if(a==='copy1')App.copy(window.__P1,'Root cause prompt copied');
else if(a==='copy2')App.copy(window.__P2,'ServiceNow prompt copied');
else if(a==='copy3')App.copy(window.__P3,'Plain English prompt copied');
else if(a==='copyMsg')App.copy(Plain.message(),'Status message copied - paste into Zoom, Teams or email');
else if(a==='clearLedger'){Ledger.clear();toast('Capture ledger cleared','warn');setTimeout(function(){location.reload();},600);}});
document.addEventListener('mouseover',function(e){var t=e.target.closest('.cmd-item');if(!t)return;Array.prototype.slice.call(document.querySelectorAll('#cmdList .cmd-item')).forEach(function(x){x.classList.remove('sel');});t.classList.add('sel');});
q('#cmdInput').addEventListener('input',function(){Palette.list(this.value);});
Array.prototype.slice.call(document.querySelectorAll('#rangePills .pill')).forEach(function(p){p.addEventListener('click',function(){App.setWin(parseInt(p.getAttribute('data-win'),10));});});
document.addEventListener('keydown',function(e){
var pal=q('#palette').classList.contains('show');
if((e.metaKey||e.ctrlKey)&&e.key.toLowerCase()==='k'){e.preventDefault();if(pal){Palette.close();}else{Palette.open();}return;}
if(pal){if(e.key==='Escape')Palette.close();if(e.key==='ArrowDown'){e.preventDefault();Palette.move(1);}if(e.key==='ArrowUp'){e.preventDefault();Palette.move(-1);}if(e.key==='Enter'){e.preventDefault();Palette.runSel();}return;}
if(e.target.tagName==='INPUT')return;
var k=e.key.toLowerCase();
if(e.key==='Escape')App.closeHelp();
else if(k==='x')Panels.focus('explain');
else if(k==='t')App.theme();
else if(k==='f')App.full();
else if(k==='r')Panels.renderAll();
else if(k==='e')App.exportJson();
else if(k==='p')window.print();
else if(k==='c')App.copy(window.__P1,'Root cause prompt copied');
else if(k==='m')App.copy(Plain.message(),'Status message copied');
else if(k==='?')App.help();
else if(k==='1')App.setWin(25);
else if(k==='2')App.setWin(50);
else if(k==='3')App.setWin(100);
else if(k==='4')App.setWin(0);});
window.addEventListener('resize',function(){Panels.renderAll();});
App.init();
'@

# ------------------------------------------------------------ write report ---

$head = "<!DOCTYPE html><html lang='en'><head><meta charset='UTF-8'>" +
        "<meta name='viewport' content='width=device-width, initial-scale=1.0, viewport-fit=cover'>" +
        "<meta name='theme-color' content='#05080d'>" +
        "<title>WiFi Connectivity Monitor - $hostName</title><style>"

$doc = New-Object System.Text.StringBuilder
[void]$doc.Append($head)
[void]$doc.Append($cssText)
[void]$doc.Append($shellText)
[void]$doc.Append("var CAP=$capJs;`r`n")
[void]$doc.Append($jsText)
[void]$doc.Append("`r`n</scr" + "ipt></body></html>")

$utf8 = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($htmlPath, $doc.ToString(), $utf8)

Write-Log "Report written: $htmlPath"

try { Start-Process $htmlPath } catch { Write-Host "Open the report manually: $htmlPath" }

$sigLine = "not available"
if ($haveSig) { $sigLine = "$rssiVal dBm (derived from $sigPct% link quality)" }

$summary = @"
WiFi Connectivity Monitor v$TOOL_VERSION complete.

Capture point:  $Location
Link:           $linkType
Packet loss:    $lossPct percent
Avg latency:    $avgLat ms
Jitter:         $jitterVal ms
Longest drop:   $maxConsecFail samples
Signal:         $sigLine
Gateway:        $gwStatus

Tip: open the PLAIN ENGLISH BRIEFING panel (or press X) for a
non-technical explanation and a ready-to-send status message.

Report folder:
$outDir
"@

Write-Host ""
Write-Host $summary -ForegroundColor Cyan
Write-Host ""

if ($hasForms) {
    [void][System.Windows.Forms.MessageBox]::Show($summary, "WiFi Connectivity Monitor v$TOOL_VERSION",
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}
