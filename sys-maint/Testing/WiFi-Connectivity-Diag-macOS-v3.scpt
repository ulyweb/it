-- ============================================================================
-- WIFI CONNECTIVITY MONITOR v3.0
-- 
-- 
--
-- Interface pattern modeled on worldmonitor.app / PG&E POWER MONITOR v4.0:
--   sticky topbar, LIVE pulse, version + status pills, command palette (Cmd K),
--   keyboard shortcuts, dark/light theme, panel chips, collapsible panels,
--   KPI strip, zero-dependency canvas charts, toasts, JSON export, print/PDF,
--   and a localStorage capture ledger for multi-position comparison.
--
-- READ-ONLY capture. No network settings are changed.
-- Save As: Script Editor -> File -> Save -> File Format: Application
-- Output:  /Users/Shared/WiFiConnectivityDiag/reports/<timestamp>_<location>/
--
-- Every string literal in this script is single-line by design so that it
-- always compiles cleanly in Script Editor.
-- ============================================================================

property PING_TARGET : "8.8.8.8"
property DNS_TARGET : "google.com"
property HTTP_TARGET : "https://www.apple.com/library/test/success.html"
property BASE_DIR : "/Users/Shared/WiFiConnectivityDiag"
property TOOL_VERSION : "3.0"

on strReplace(theText, searchStr, replaceStr)
	set od to AppleScript's text item delimiters
	set AppleScript's text item delimiters to searchStr
	set theParts to text items of theText
	set AppleScript's text item delimiters to replaceStr
	set newText to theParts as text
	set AppleScript's text item delimiters to od
	return newText
end strReplace

on trimText(theText)
	set s to theText
	repeat while s starts with " " or s starts with tab
		if (length of s) is less than 2 then
			set s to ""
			exit repeat
		end if
		set s to text 2 thru -1 of s
	end repeat
	repeat while s ends with " " or s ends with tab
		if (length of s) is less than 2 then
			set s to ""
			exit repeat
		end if
		set s to text 1 thru -2 of s
	end repeat
	return s
end trimText

on jsEsc(theText)
	set bs to "\\"
	set t to my strReplace(theText, bs, bs & bs)
	set t to my strReplace(t, "'", bs & "'")
	set t to my strReplace(t, return, "")
	set t to my strReplace(t, linefeed, bs & "n")
	set t to my strReplace(t, tab, "  ")
	return t
end jsEsc

on runShell(cmdText)
	try
		return do shell script cmdText
	on error
		return ""
	end try
end runShell

on runShellAll(cmdText)
	try
		return do shell script cmdText & " 2>&1"
	on error errMsg
		return "Command did not complete: " & errMsg
	end try
end runShellAll

on writeNew(filePath, theText)
	do shell script "printf '%s' " & quoted form of theText & " > " & quoted form of filePath
end writeNew

on appendText(filePath, theText)
	do shell script "printf '%s' " & quoted form of theText & " >> " & quoted form of filePath
end appendText

on valueForKey(blockText, keyName)
	repeat with p in paragraphs of blockText
		set theLine to p as text
		if theLine contains (keyName & ":") then
			set od to AppleScript's text item delimiters
			set AppleScript's text item delimiters to (keyName & ":")
			set theParts to text items of theLine
			set AppleScript's text item delimiters to od
			if (count of theParts) is greater than 1 then
				return my trimText(item 2 of theParts)
			end if
		end if
	end repeat
	return ""
end valueForKey

on firstField(theText, sepStr)
	set od to AppleScript's text item delimiters
	set AppleScript's text item delimiters to sepStr
	set theParts to text items of theText
	set AppleScript's text item delimiters to od
	if (count of theParts) is 0 then return ""
	return item 1 of theParts
end firstField

on roundTo2(theValue)
	return ((round (theValue * 100)) / 100)
end roundTo2

-- ------------------------------------------------------- capture options ---

set locOptions to {"Desk / User Location", "Near Router or Access Point", "Wired Ethernet Test", "Conference Room", "Reception or Lobby", "Other Location"}
set locPick to choose from list locOptions with title "WiFi Connectivity Monitor v3.0" with prompt "Where is this capture being taken? Captures are stored and compared against each other in the dashboard." default items {"Desk / User Location"}
if locPick is false then return
set captureLocation to item 1 of locPick

set modeOptions to {"Quick - 20 samples (about 20 sec)", "Standard - 60 samples (about 1 min)", "Extended - 180 samples (about 3 min)", "Deep - 300 samples (about 5 min)"}
set modePick to choose from list modeOptions with title "WiFi Connectivity Monitor v3.0" with prompt "Select stability test length. Longer captures catch intermittent drops far more reliably." default items {"Standard - 60 samples (about 1 min)"}
if modePick is false then return
set modeChoice to item 1 of modePick

set pingCount to 60
if modeChoice starts with "Quick" then set pingCount to 20
if modeChoice starts with "Extended" then set pingCount to 180
if modeChoice starts with "Deep" then set pingCount to 300

-- --------------------------------------------------------- output folders ---

set stampText to do shell script "date +%Y%m%d_%H%M%S"
set safeLabel to do shell script "echo " & quoted form of captureLocation & " | tr ' /' '__' | tr -cd 'A-Za-z0-9_'"
set outDir to BASE_DIR & "/reports/" & stampText & "_" & safeLabel
do shell script "mkdir -p " & quoted form of outDir

set htmlPath to outDir & "/WiFi_Connectivity_Monitor.html"
set csvPath to outDir & "/PingSamples.csv"
set logPath to outDir & "/CaptureLog.log"
set spAirPath to outDir & "/SPAirPortDataType.txt"

my writeNew(logPath, "[" & (do shell script "date '+%Y-%m-%d %H:%M:%S'") & "] WiFi Connectivity Monitor v3.0 started" & linefeed)
my appendText(logPath, "Location: " & captureLocation & "  Samples: " & (pingCount as text) & linefeed)

-- ---------------------------------------------------------- device details ---

set hostName to my runShell("scutil --get ComputerName || hostname")
set consoleUser to my runShell("stat -f %Su /dev/console")
set hwBlock to my runShellAll("system_profiler SPHardwareDataType")
set modelName to my valueForKey(hwBlock, "Model Name")
set modelID to my valueForKey(hwBlock, "Model Identifier")
set chipName to my valueForKey(hwBlock, "Chip")
if chipName is "" then set chipName to my valueForKey(hwBlock, "Processor Name")
set serialNum to my valueForKey(hwBlock, "Serial Number (system)")
if serialNum is "" then set serialNum to my valueForKey(hwBlock, "Serial Number")
set osVer to my runShell("sw_vers -productVersion")
set osBuild to my runShell("sw_vers -buildVersion")
set osName to my runShell("sw_vers -productName")
if osName is "" then set osName to "macOS"

-- --------------------------------------------------------------- interfaces ---

set wifiDevice to my runShell("networksetup -listallhardwareports | awk '/Wi-Fi|AirPort/{getline; print $2; exit}'")
if wifiDevice is "" then set wifiDevice to "en0"
set defaultIface to my runShell("route -n get default 2>/dev/null | awk '/interface:/{print $2; exit}'")
if defaultIface is "" then set defaultIface to wifiDevice
set routerIP to my runShell("route -n get default 2>/dev/null | awk '/gateway:/{print $2; exit}'")
set activeIP to my runShell("ipconfig getifaddr " & defaultIface)
if activeIP is "" then set activeIP to my runShell("ipconfig getifaddr " & wifiDevice)
set dnsServers to my runShell("scutil --dns 2>/dev/null | awk '/nameserver\\[[0-9]+\\]/{print $3}' | sort -u | paste -sd ', ' -")

set linkType to "Wi-Fi"
if defaultIface is not wifiDevice then set linkType to "Wired or Other"
if routerIP is "" then set routerIP to "Not detected"
if activeIP is "" then set activeIP to "Not detected"
if dnsServers is "" then set dnsServers to "Not detected"

-- ------------------------------------------------------ wireless telemetry ---

set spAir to my runShellAll("system_profiler SPAirPortDataType")
my writeNew(spAirPath, spAir & linefeed)

set ssidName to my runShell("awk '/Current Network Information:/{getline; print; exit}' " & quoted form of spAirPath & " | sed 's/^[[:space:]]*//; s/[[:space:]]*:[[:space:]]*$//'")
if ssidName is "" then set ssidName to my runShell("ipconfig getsummary " & wifiDevice & " 2>/dev/null | awk -F' SSID : ' '/ SSID : /{print $2; exit}'")
if ssidName is "" then set ssidName to "Not available"

set phyMode to my valueForKey(spAir, "PHY Mode")
set wChannel to my valueForKey(spAir, "Channel")
set wSecurity to my valueForKey(spAir, "Security")
set txRate to my valueForKey(spAir, "Transmit Rate")
set bssidVal to my valueForKey(spAir, "BSSID")
set sigNoise to my valueForKey(spAir, "Signal / Noise")

set rssiVal to 0
set noiseVal to 0
set snrVal to 0
set haveSignal to false
if sigNoise is not "" then
	set od to AppleScript's text item delimiters
	set AppleScript's text item delimiters to " / "
	set snParts to text items of sigNoise
	set AppleScript's text item delimiters to od
	if (count of snParts) is greater than 1 then
		try
			set rssiVal to (my firstField(my trimText(item 1 of snParts), " ")) as integer
			set noiseVal to (my firstField(my trimText(item 2 of snParts), " ")) as integer
			set snrVal to rssiVal - noiseVal
			set haveSignal to true
		end try
	end if
end if

if phyMode is "" then set phyMode to "N/A"
if wChannel is "" then set wChannel to "N/A"
if wSecurity is "" then set wSecurity to "N/A"
if txRate is "" then set txRate to "N/A"
if bssidVal is "" then set bssidVal to "N/A"

set neighborTotal to my runShell("awk '/Other Local Wi-Fi Networks:/{f=1} f&&/Channel:/{c++} END{print c+0}' " & quoted form of spAirPath)
if neighborTotal is "" then set neighborTotal to "0"
set neighborRaw to my runShellAll("awk -F': ' '/Other Local Wi-Fi Networks:/{f=1} f&&/Channel:/{print $2}' " & quoted form of spAirPath & " | sed 's/[^0-9].*//' | grep -v '^$' | sort -n | uniq -c | sort -rn | head -20")

set neighborJs to ""
try
	repeat with p in paragraphs of neighborRaw
		set theLine to my trimText(p as text)
		if theLine is not "" then
			set nCount to my firstField(theLine, " ")
			set od to AppleScript's text item delimiters
			set AppleScript's text item delimiters to " "
			set bits to text items of theLine
			set AppleScript's text item delimiters to od
			if (count of bits) is greater than 1 then
				set chVal to item -1 of bits
				try
					set nInt to nCount as integer
					set cInt to chVal as integer
					set neighborJs to neighborJs & "{ch:" & (cInt as text) & ",n:" & (nInt as text) & "},"
				end try
			end if
		end if
	end repeat
end try

-- ---------------------------------------------------------- stability test ---

set csvText to "Sample,Timestamp,Target,Status,LatencyMs" & linefeed
set samplesJs to ""
set successCount to 0
set failCount to 0
set latSum to 0.0
set latMax to 0.0
set latMin to 999999.0
set consecFail to 0
set maxConsecFail to 0
set prevLat to -1.0
set jitterSum to 0.0
set jitterCount to 0

set progress total steps to pingCount
set progress completed steps to 0
set progress description to "Capturing network telemetry..."
set progress additional description to "Target 8.8.8.8"

repeat with i from 1 to pingCount
	set clockNow to do shell script "date '+%H:%M:%S'"
	set stampNow to do shell script "date '+%Y-%m-%d %H:%M:%S'"
	set latRaw to ""
	try
		set latRaw to do shell script "ping -c 1 -W 1000 " & PING_TARGET & " 2>/dev/null | awk -F'time=' '/time=/{printf \"%d\", $2*1000; exit}'"
	end try
	if latRaw is "" then
		set failCount to failCount + 1
		set consecFail to consecFail + 1
		if consecFail is greater than maxConsecFail then set maxConsecFail to consecFail
		set csvText to csvText & (i as text) & "," & stampNow & "," & PING_TARGET & ",Failed," & linefeed
		set samplesJs to samplesJs & "{i:" & (i as text) & ",t:'" & clockNow & "',ok:0,ms:0},"
	else
		set consecFail to 0
		set successCount to successCount + 1
		set latVal to ((latRaw as integer) / 1000.0)
		set latSum to latSum + latVal
		if latVal is greater than latMax then set latMax to latVal
		if latVal is less than latMin then set latMin to latVal
		if prevLat is greater than 0 then
			set d to latVal - prevLat
			if d is less than 0 then set d to -d
			set jitterSum to jitterSum + d
			set jitterCount to jitterCount + 1
		end if
		set prevLat to latVal
		set csvText to csvText & (i as text) & "," & stampNow & "," & PING_TARGET & ",Success," & ((my roundTo2(latVal)) as text) & linefeed
		set samplesJs to samplesJs & "{i:" & (i as text) & ",t:'" & clockNow & "',ok:1,ms:" & ((my roundTo2(latVal)) as text) & "},"
	end if
	set progress completed steps to i
	set progress additional description to "Sample " & (i as text) & " of " & (pingCount as text)
	delay 1
end repeat

set progress completed steps to pingCount
my writeNew(csvPath, csvText)

set lossPct to my roundTo2((failCount / pingCount) * 100)
set avgLat to 0
if successCount is greater than 0 then set avgLat to my roundTo2(latSum / successCount)
set maxLat to 0
set minLat to 0
if successCount is greater than 0 then
	set maxLat to my roundTo2(latMax)
	set minLat to my roundTo2(latMin)
end if
set jitterVal to 0
if jitterCount is greater than 0 then set jitterVal to my roundTo2(jitterSum / jitterCount)

-- --------------------------------------------------------- path validation ---

set gwStatus to "unknown"
set gwTest to "Default gateway was not detected."
if routerIP is not "Not detected" then
	set gwTest to my runShellAll("ping -c 4 -W 1000 " & routerIP & " | tail -4")
	if gwTest contains "0.0% packet loss" then
		set gwStatus to "clean"
	else if gwTest contains "packet loss" then
		set gwStatus to "lossy"
	else
		set gwStatus to "unreachable"
	end if
end if

set dnsTest to my runShellAll("dig +time=2 +tries=1 " & DNS_TARGET & " 2>/dev/null || nslookup " & DNS_TARGET)
set dnsOk to "0"
if dnsTest contains "ANSWER SECTION" then set dnsOk to "1"
if dnsTest contains "Address" then set dnsOk to "1"

set httpTest to my runShellAll("curl -s -o /dev/null -w 'dns=%{time_namelookup}s tcp=%{time_connect}s tls=%{time_appconnect}s total=%{time_total}s code=%{http_code}' --max-time 15 " & HTTP_TARGET)
set httpCode to my runShell("curl -s -o /dev/null -w '%{http_code}' --max-time 15 " & HTTP_TARGET)
if httpCode is "" then set httpCode to "000"

set traceTest to my runShellAll("traceroute -n -w 1 -q 1 -m 12 " & PING_TARGET)
set hop1Loss to "0"
try
	set hop1Line to my runShell("traceroute -n -w 1 -q 2 -m 1 " & PING_TARGET & " 2>/dev/null | sed -n '2p'")
	if hop1Line contains "*" then set hop1Loss to "1"
end try

set routeTable to my runShellAll("netstat -rn | head -26")
set ifconfigOut to my runShellAll("ifconfig " & wifiDevice)
set scutilDNS to my runShellAll("scutil --dns | head -36")
set preferredNets to my runShellAll("networksetup -listpreferredwirelessnetworks " & wifiDevice)

my appendText(logPath, "Loss " & (lossPct as text) & " pct  Avg " & (avgLat as text) & " ms  Jitter " & (jitterVal as text) & " ms  MaxDrop " & (maxConsecFail as text) & linefeed)

-- ------------------------------------------------------------ data payload ---

set epochMs to (do shell script "date +%s") & "000"

set capJs to "{"
set capJs to capJs & "v:'3.0',"
set capJs to capJs & "id:'" & stampText & "_" & safeLabel & "',"
set capJs to capJs & "ts:" & epochMs & ","
set capJs to capJs & "loc:'" & my jsEsc(captureLocation) & "',"
set capJs to capJs & "host:'" & my jsEsc(hostName) & "',"
set capJs to capJs & "user:'" & my jsEsc(consoleUser) & "',"
set capJs to capJs & "model:'" & my jsEsc(modelName) & "',"
set capJs to capJs & "modelId:'" & my jsEsc(modelID) & "',"
set capJs to capJs & "chip:'" & my jsEsc(chipName) & "',"
set capJs to capJs & "serial:'" & my jsEsc(serialNum) & "',"
set capJs to capJs & "osName:'" & my jsEsc(osName) & "',"
set capJs to capJs & "osVer:'" & my jsEsc(osVer) & "',"
set capJs to capJs & "osBuild:'" & my jsEsc(osBuild) & "',"
set capJs to capJs & "wifiDev:'" & my jsEsc(wifiDevice) & "',"
set capJs to capJs & "iface:'" & my jsEsc(defaultIface) & "',"
set capJs to capJs & "link:'" & my jsEsc(linkType) & "',"
set capJs to capJs & "ip:'" & my jsEsc(activeIP) & "',"
set capJs to capJs & "router:'" & my jsEsc(routerIP) & "',"
set capJs to capJs & "dns:'" & my jsEsc(dnsServers) & "',"
set capJs to capJs & "ssid:'" & my jsEsc(ssidName) & "',"
set capJs to capJs & "bssid:'" & my jsEsc(bssidVal) & "',"
set capJs to capJs & "phy:'" & my jsEsc(phyMode) & "',"
set capJs to capJs & "chan:'" & my jsEsc(wChannel) & "',"
set capJs to capJs & "sec:'" & my jsEsc(wSecurity) & "',"
set capJs to capJs & "tx:'" & my jsEsc(txRate) & "',"
set capJs to capJs & "haveSig:" & (haveSignal as text) & ","
set capJs to capJs & "rssi:" & (rssiVal as text) & ","
set capJs to capJs & "noise:" & (noiseVal as text) & ","
set capJs to capJs & "snr:" & (snrVal as text) & ","
set capJs to capJs & "nbTotal:" & neighborTotal & ","
set capJs to capJs & "nb:[" & neighborJs & "],"
set capJs to capJs & "target:'" & PING_TARGET & "',"
set capJs to capJs & "n:" & (pingCount as text) & ","
set capJs to capJs & "ok:" & (successCount as text) & ","
set capJs to capJs & "bad:" & (failCount as text) & ","
set capJs to capJs & "loss:" & (lossPct as text) & ","
set capJs to capJs & "avg:" & (avgLat as text) & ","
set capJs to capJs & "min:" & (minLat as text) & ","
set capJs to capJs & "max:" & (maxLat as text) & ","
set capJs to capJs & "jit:" & (jitterVal as text) & ","
set capJs to capJs & "drop:" & (maxConsecFail as text) & ","
set capJs to capJs & "gw:'" & gwStatus & "',"
set capJs to capJs & "dnsOk:" & dnsOk & ","
set capJs to capJs & "httpCode:'" & my jsEsc(httpCode) & "',"
set capJs to capJs & "hop1:" & hop1Loss & ","
set capJs to capJs & "samples:[" & samplesJs & "],"
set capJs to capJs & "raw:{"
set capJs to capJs & "gw:'" & my jsEsc(gwTest) & "',"
set capJs to capJs & "dns:'" & my jsEsc(dnsTest) & "',"
set capJs to capJs & "http:'" & my jsEsc(httpTest) & "',"
set capJs to capJs & "trace:'" & my jsEsc(traceTest) & "',"
set capJs to capJs & "route:'" & my jsEsc(routeTable) & "',"
set capJs to capJs & "ifc:'" & my jsEsc(ifconfigOut) & "',"
set capJs to capJs & "scutil:'" & my jsEsc(scutilDNS) & "',"
set capJs to capJs & "pref:'" & my jsEsc(preferredNets) & "',"
set capJs to capJs & "air:'" & my jsEsc(spAir) & "'},"
set capJs to capJs & "out:'" & my jsEsc(outDir) & "'}"

-- ------------------------------------------------------------- stylesheet ---

set cssText to ""

set cssText to cssText & "*,*::before,*::after{box-sizing:border-box}" & linefeed
set cssText to cssText & ":root{--bg:#05080d;--bg2:#080d15;--grid-line:rgba(0,229,255,.045);--panel:rgba(11,17,27,.86);--panel-solid:#0b111b;--border:rgba(0,229,255,.16);--border-soft:rgba(255,255,255,.07);--txt:#d7e6f2;--txt-dim:#7d93a8;--txt-mute:#4e6275;--cyan:#00e5ff;--green:#00e08a;--amber:#ffb020;--red:#ff3b57;--violet:#a97bff;--mono:'SFMono-Regular',ui-monospace,Menlo,Consolas,monospace;--sans:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;--r:10px;--shadow:0 10px 34px rgba(0,0,0,.6)}" & linefeed
set cssText to cssText & "html[data-theme='light']{--bg:#eef2f7;--bg2:#e4eaf2;--grid-line:rgba(0,80,120,.05);--panel:rgba(255,255,255,.92);--panel-solid:#fff;--border:rgba(0,110,150,.22);--border-soft:rgba(0,0,0,.09);--txt:#0f1c28;--txt-dim:#456173;--txt-mute:#7b93a5;--cyan:#0092b8;--green:#00875a;--amber:#b06f00;--red:#c62a42;--violet:#6b3fd4;--shadow:0 8px 26px rgba(20,40,60,.14)}" & linefeed
set cssText to cssText & "html,body{height:100%}" & linefeed
set cssText to cssText & "body{margin:0;background:var(--bg);color:var(--txt);font-family:var(--sans);font-size:14px;line-height:1.5;-webkit-font-smoothing:antialiased;overflow-x:hidden;background-image:linear-gradient(var(--grid-line) 1px,transparent 1px),linear-gradient(90deg,var(--grid-line) 1px,transparent 1px);background-size:46px 46px}" & linefeed
set cssText to cssText & "body::before{content:'';position:fixed;inset:0;pointer-events:none;z-index:0;background:radial-gradient(1100px 620px at 12% -8%,rgba(0,229,255,.10),transparent 62%),radial-gradient(900px 520px at 92% 4%,rgba(169,123,255,.08),transparent 60%)}" & linefeed
set cssText to cssText & "::-webkit-scrollbar{width:9px;height:9px}" & linefeed
set cssText to cssText & "::-webkit-scrollbar-track{background:transparent}" & linefeed
set cssText to cssText & "::-webkit-scrollbar-thumb{background:rgba(0,229,255,.22);border-radius:9px}" & linefeed
set cssText to cssText & "::-webkit-scrollbar-thumb:hover{background:rgba(0,229,255,.4)}" & linefeed
set cssText to cssText & "button,input,select{font-family:inherit;color:inherit}" & linefeed
set cssText to cssText & "button{cursor:pointer}" & linefeed
set cssText to cssText & "a{color:var(--cyan)}" & linefeed
set cssText to cssText & ".topbar{position:sticky;top:0;z-index:60;display:flex;align-items:center;gap:10px;flex-wrap:wrap;padding:8px 14px;background:linear-gradient(180deg,rgba(5,8,13,.97),rgba(5,8,13,.86));border-bottom:1px solid var(--border);backdrop-filter:blur(14px);-webkit-backdrop-filter:blur(14px)}" & linefeed
set cssText to cssText & "html[data-theme='light'] .topbar{background:linear-gradient(180deg,rgba(255,255,255,.97),rgba(255,255,255,.86))}" & linefeed
set cssText to cssText & ".brand{display:flex;align-items:center;gap:8px;font-family:var(--mono);font-weight:700;letter-spacing:.13em;font-size:13px;white-space:nowrap}" & linefeed
set cssText to cssText & ".brand .bolt{font-size:17px;filter:drop-shadow(0 0 8px var(--cyan))}" & linefeed
set cssText to cssText & ".brand .b1{color:var(--cyan)}" & linefeed
set cssText to cssText & ".brand .b2{color:var(--txt)}" & linefeed
set cssText to cssText & ".ver{font-size:9px;color:var(--txt-mute);border:1px solid var(--border-soft);padding:1px 5px;border-radius:4px;font-family:var(--mono)}" & linefeed
set cssText to cssText & ".live{display:inline-flex;align-items:center;gap:5px;font-family:var(--mono);font-size:10px;letter-spacing:.14em;color:var(--green);border:1px solid rgba(0,224,138,.34);background:rgba(0,224,138,.08);padding:2px 7px;border-radius:20px}" & linefeed
set cssText to cssText & ".live .dot{width:6px;height:6px;border-radius:50%;background:var(--green);animation:blip 1.6s infinite}" & linefeed
set cssText to cssText & "@keyframes blip{0%,100%{opacity:1;box-shadow:0 0 0 0 rgba(0,224,138,.6)}55%{opacity:.35;box-shadow:0 0 0 7px rgba(0,224,138,0)}}" & linefeed
set cssText to cssText & ".tb-spacer{flex:1 1 auto}" & linefeed
set cssText to cssText & ".clockbox{font-family:var(--mono);font-size:11px;color:var(--txt-dim);display:flex;flex-direction:column;line-height:1.28;text-align:right}" & linefeed
set cssText to cssText & ".clockbox b{color:var(--txt);font-size:13px;letter-spacing:.05em}" & linefeed
set cssText to cssText & ".tb-btn{background:rgba(255,255,255,.035);border:1px solid var(--border-soft);color:var(--txt-dim);border-radius:7px;padding:5px 9px;font-size:11px;font-family:var(--mono);letter-spacing:.06em;display:inline-flex;align-items:center;gap:5px;transition:.16s;white-space:nowrap}" & linefeed
set cssText to cssText & ".tb-btn:hover{border-color:var(--cyan);color:var(--cyan);background:rgba(0,229,255,.09);transform:translateY(-1px)}" & linefeed
set cssText to cssText & ".kbd{font-family:var(--mono);font-size:9px;border:1px solid var(--border-soft);border-radius:3px;padding:0 4px;color:var(--txt-mute);background:rgba(255,255,255,.04)}" & linefeed
set cssText to cssText & ".searchbtn{min-width:160px;justify-content:space-between}" & linefeed
set cssText to cssText & ".subbar{position:sticky;top:46px;z-index:55;display:flex;align-items:center;gap:6px;overflow-x:auto;padding:7px 14px;background:rgba(5,8,13,.9);border-bottom:1px solid var(--border-soft);backdrop-filter:blur(10px);scrollbar-width:none}" & linefeed
set cssText to cssText & "html[data-theme='light'] .subbar{background:rgba(255,255,255,.9)}" & linefeed
set cssText to cssText & ".subbar::-webkit-scrollbar{display:none}" & linefeed
set cssText to cssText & ".sb-label{font-family:var(--mono);font-size:9px;letter-spacing:.16em;color:var(--txt-mute);white-space:nowrap;padding-right:2px}" & linefeed
set cssText to cssText & ".chip{font-family:var(--mono);font-size:10px;letter-spacing:.05em;white-space:nowrap;border:1px solid var(--border-soft);background:rgba(255,255,255,.03);color:var(--txt-mute);padding:3px 9px;border-radius:20px;transition:.16s}" & linefeed
set cssText to cssText & ".chip:hover{color:var(--txt);border-color:var(--border)}" & linefeed
set cssText to cssText & ".chip.on{color:var(--cyan);border-color:rgba(0,229,255,.45);background:rgba(0,229,255,.12);box-shadow:0 0 12px rgba(0,229,255,.14) inset}" & linefeed
set cssText to cssText & ".range-pills{display:flex;gap:3px;margin-left:auto;align-items:center}" & linefeed
set cssText to cssText & ".pill{font-family:var(--mono);font-size:10px;padding:3px 8px;border-radius:5px;border:1px solid var(--border-soft);background:transparent;color:var(--txt-mute)}" & linefeed
set cssText to cssText & ".pill.on{background:rgba(0,229,255,.16);border-color:var(--cyan);color:var(--cyan)}" & linefeed
set cssText to cssText & ".wrap{position:relative;z-index:1;padding:14px;max-width:1680px;margin:0 auto}" & linefeed
set cssText to cssText & "#alertBanner{display:none;align-items:center;gap:10px;margin-bottom:12px;padding:11px 15px;border-radius:var(--r);border:1px solid rgba(255,59,87,.5);background:linear-gradient(90deg,rgba(255,59,87,.2),rgba(255,59,87,.05));font-family:var(--mono);font-size:12px;letter-spacing:.06em;color:#ffd7dd;animation:pulseBar 2.2s infinite}" & linefeed
set cssText to cssText & "@keyframes pulseBar{0%,100%{box-shadow:0 0 0 0 rgba(255,59,87,.35)}50%{box-shadow:0 0 22px 2px rgba(255,59,87,.14)}}" & linefeed
set cssText to cssText & "#alertBanner.show{display:flex}" & linefeed
set cssText to cssText & "#alertBanner .x{margin-left:auto;background:none;border:none;color:#ffd7dd;font-size:14px}" & linefeed
set cssText to cssText & ".hero{border:1px solid var(--border);border-radius:14px;background:var(--panel);backdrop-filter:blur(12px);box-shadow:var(--shadow);overflow:hidden;margin-bottom:14px;position:relative}" & linefeed
set cssText to cssText & ".hero::after{content:'';position:absolute;left:0;right:0;top:0;height:2px;background:linear-gradient(90deg,transparent,var(--cyan),var(--violet),transparent);background-size:200% 100%;animation:shimmer 4.5s linear infinite}" & linefeed
set cssText to cssText & "@keyframes shimmer{0%{background-position:200% 0}100%{background-position:-200% 0}}" & linefeed
set cssText to cssText & ".hero-top{display:flex;flex-wrap:wrap;align-items:center;gap:14px;padding:15px 18px;border-bottom:1px solid var(--border-soft)}" & linefeed
set cssText to cssText & ".statuslamp{width:62px;height:62px;border-radius:50%;display:grid;place-items:center;font-size:27px;flex:none;position:relative}" & linefeed
set cssText to cssText & ".statuslamp::after{content:'';position:absolute;inset:-6px;border-radius:50%;border:2px solid currentColor;opacity:.32;animation:ring 2.4s infinite}" & linefeed
set cssText to cssText & "@keyframes ring{0%{transform:scale(.86);opacity:.55}100%{transform:scale(1.22);opacity:0}}" & linefeed
set cssText to cssText & ".lamp-ok{background:radial-gradient(circle at 35% 30%,#00e08a,#046b45);color:var(--green);box-shadow:0 0 34px rgba(0,224,138,.5)}" & linefeed
set cssText to cssText & ".lamp-warn{background:radial-gradient(circle at 35% 30%,#ffc857,#8a5b00);color:var(--amber);box-shadow:0 0 34px rgba(255,176,32,.5)}" & linefeed
set cssText to cssText & ".lamp-bad{background:radial-gradient(circle at 35% 30%,#ff6076,#8d0d22);color:var(--red);box-shadow:0 0 34px rgba(255,59,87,.5)}" & linefeed
set cssText to cssText & ".hero-id{flex:1 1 250px;min-width:210px}" & linefeed
set cssText to cssText & ".hero-id .lbl{font-family:var(--mono);font-size:9.5px;letter-spacing:.2em;color:var(--txt-mute)}" & linefeed
set cssText to cssText & ".hero-id h1{margin:2px 0 3px;font-size:25px;font-weight:750;letter-spacing:-.4px}" & linefeed
set cssText to cssText & ".hero-id .sub{font-family:var(--mono);font-size:11px;color:var(--txt-dim)}" & linefeed
set cssText to cssText & ".hero-rate{text-align:right;flex:none}" & linefeed
set cssText to cssText & ".hero-rate .big{font-family:var(--mono);font-size:40px;font-weight:750;line-height:1;letter-spacing:-1.5px}" & linefeed
set cssText to cssText & ".hero-rate .unit{font-family:var(--mono);font-size:10px;color:var(--txt-mute);letter-spacing:.16em}" & linefeed
set cssText to cssText & ".statuspill{display:inline-flex;align-items:center;gap:5px;font-family:var(--mono);font-size:10px;letter-spacing:.14em;padding:3px 9px;border-radius:20px;border:1px solid}" & linefeed
set cssText to cssText & ".sp-ok{color:var(--green);border-color:rgba(0,224,138,.45);background:rgba(0,224,138,.1)}" & linefeed
set cssText to cssText & ".sp-warn{color:var(--amber);border-color:rgba(255,176,32,.45);background:rgba(255,176,32,.1)}" & linefeed
set cssText to cssText & ".sp-bad{color:var(--red);border-color:rgba(255,59,87,.45);background:rgba(255,59,87,.1)}" & linefeed
set cssText to cssText & ".kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(148px,1fr));gap:1px;background:var(--border-soft)}" & linefeed
set cssText to cssText & ".kpi{background:var(--panel-solid);padding:11px 14px;transition:.18s}" & linefeed
set cssText to cssText & ".kpi:hover{background:rgba(0,229,255,.06);transform:translateY(-2px)}" & linefeed
set cssText to cssText & ".kpi .k{font-family:var(--mono);font-size:9px;letter-spacing:.15em;color:var(--txt-mute);margin-bottom:3px}" & linefeed
set cssText to cssText & ".kpi .v{font-family:var(--mono);font-size:19px;font-weight:700;letter-spacing:-.5px}" & linefeed
set cssText to cssText & ".kpi .d{font-family:var(--mono);font-size:9.5px;color:var(--txt-dim);margin-top:1px}" & linefeed
set cssText to cssText & ".grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(360px,1fr));gap:13px;align-items:start}" & linefeed
set cssText to cssText & ".panel{border:1px solid var(--border-soft);border-radius:var(--r);background:var(--panel);backdrop-filter:blur(10px);box-shadow:var(--shadow);overflow:hidden;animation:fadeUp .42s cubic-bezier(.2,.7,.3,1) both;position:relative}" & linefeed
set cssText to cssText & "@keyframes fadeUp{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:none}}" & linefeed
set cssText to cssText & ".panel.wide{grid-column:span 2}" & linefeed
set cssText to cssText & ".panel.glow{border-color:rgba(0,229,255,.5);box-shadow:0 0 26px rgba(0,229,255,.16)}" & linefeed
set cssText to cssText & ".p-head{display:flex;align-items:center;gap:8px;padding:9px 12px;border-bottom:1px solid var(--border-soft);background:linear-gradient(180deg,rgba(0,229,255,.05),transparent);cursor:pointer;user-select:none}" & linefeed
set cssText to cssText & ".p-head h3{margin:0;font-family:var(--mono);font-size:11px;letter-spacing:.14em;font-weight:600;color:var(--txt);display:flex;align-items:center;gap:6px;flex:1;min-width:0}" & linefeed
set cssText to cssText & ".badge{font-family:var(--mono);font-size:9px;padding:1px 6px;border-radius:20px;border:1px solid var(--border-soft);color:var(--txt-mute)}" & linefeed
set cssText to cssText & ".badge.new{color:var(--cyan);border-color:rgba(0,229,255,.4);background:rgba(0,229,255,.12)}" & linefeed
set cssText to cssText & ".badge.ok{color:var(--green);border-color:rgba(0,224,138,.4);background:rgba(0,224,138,.12)}" & linefeed
set cssText to cssText & ".badge.bad{color:var(--red);border-color:rgba(255,59,87,.4);background:rgba(255,59,87,.12)}" & linefeed
set cssText to cssText & ".p-act{background:none;border:none;color:var(--txt-mute);font-size:12px;padding:1px 4px;border-radius:4px;line-height:1}" & linefeed
set cssText to cssText & ".p-act:hover{color:var(--cyan);background:rgba(0,229,255,.1)}" & linefeed
set cssText to cssText & ".p-body{padding:12px;max-height:1800px;overflow:auto;transition:max-height .3s ease,padding .3s ease}" & linefeed
set cssText to cssText & ".panel.collapsed .p-body{max-height:0;padding-top:0;padding-bottom:0;overflow:hidden}" & linefeed
set cssText to cssText & ".panel.collapsed .caret{transform:rotate(-90deg)}" & linefeed
set cssText to cssText & ".caret{transition:transform .25s;display:inline-block}" & linefeed
set cssText to cssText & ".mono{font-family:var(--mono)}" & linefeed
set cssText to cssText & ".dim{color:var(--txt-dim)}" & linefeed
set cssText to cssText & ".mute{color:var(--txt-mute)}" & linefeed
set cssText to cssText & ".c-green{color:var(--green)}" & linefeed
set cssText to cssText & ".c-amber{color:var(--amber)}" & linefeed
set cssText to cssText & ".c-red{color:var(--red)}" & linefeed
set cssText to cssText & ".c-cyan{color:var(--cyan)}" & linefeed
set cssText to cssText & ".c-violet{color:var(--violet)}" & linefeed
set cssText to cssText & ".hint{font-size:11px;color:var(--txt-mute);font-family:var(--mono);line-height:1.5}" & linefeed
set cssText to cssText & ".row{display:flex;align-items:center;gap:8px;flex-wrap:wrap}" & linefeed
set cssText to cssText & ".btn{background:rgba(0,229,255,.08);border:1px solid var(--border);color:var(--cyan);border-radius:7px;padding:7px 12px;font-family:var(--mono);font-size:11px;letter-spacing:.07em;transition:.16s}" & linefeed
set cssText to cssText & ".btn:hover{background:rgba(0,229,255,.2);transform:translateY(-1px);box-shadow:0 4px 14px rgba(0,229,255,.16)}" & linefeed
set cssText to cssText & ".btn.ghost{background:transparent;border-color:var(--border-soft);color:var(--txt-dim)}" & linefeed
set cssText to cssText & ".btn.ghost:hover{color:var(--txt);border-color:var(--border);box-shadow:none}" & linefeed
set cssText to cssText & ".btn.sm{padding:4px 8px;font-size:10px}" & linefeed
set cssText to cssText & "table{width:100%;border-collapse:collapse;font-family:var(--mono);font-size:11px}" & linefeed
set cssText to cssText & "th{text-align:left;color:var(--txt-mute);font-weight:600;letter-spacing:.1em;font-size:9px;padding:5px 6px;border-bottom:1px solid var(--border-soft);text-transform:uppercase}" & linefeed
set cssText to cssText & "td{padding:6px;border-bottom:1px solid rgba(255,255,255,.04)}" & linefeed
set cssText to cssText & "tbody tr{transition:.14s}" & linefeed
set cssText to cssText & "tbody tr:hover{background:rgba(0,229,255,.06)}" & linefeed
set cssText to cssText & "tbody tr.best td{background:rgba(0,224,138,.09)}" & linefeed
set cssText to cssText & "tbody tr.bad td{background:rgba(255,59,87,.08)}" & linefeed
set cssText to cssText & "canvas{width:100%;display:block;border-radius:7px}" & linefeed
set cssText to cssText & ".empty{text-align:center;padding:20px 8px;color:var(--txt-mute);font-family:var(--mono);font-size:11px}" & linefeed
set cssText to cssText & ".result{margin-top:10px;border:1px solid var(--border);border-radius:9px;padding:11px;background:rgba(0,229,255,.05)}" & linefeed
set cssText to cssText & ".result .line{display:flex;justify-content:space-between;font-family:var(--mono);font-size:11px;padding:2.5px 0;border-bottom:1px dashed rgba(255,255,255,.06);gap:10px}" & linefeed
set cssText to cssText & ".result .line:last-child{border:none}" & linefeed
set cssText to cssText & ".advice{margin-top:9px;padding:9px 10px;border-radius:7px;font-size:12px;line-height:1.55;border-left:3px solid}" & linefeed
set cssText to cssText & ".advice.good{background:rgba(0,224,138,.1);border-color:var(--green)}" & linefeed
set cssText to cssText & ".advice.warn{background:rgba(255,176,32,.1);border-color:var(--amber)}" & linefeed
set cssText to cssText & ".advice.bad{background:rgba(255,59,87,.1);border-color:var(--red)}" & linefeed
set cssText to cssText & ".rc{margin-bottom:9px}" & linefeed
set cssText to cssText & ".rc .rc-top{display:flex;align-items:baseline;gap:8px;font-family:var(--mono);font-size:11.5px}" & linefeed
set cssText to cssText & ".rc .rc-top b{flex:1}" & linefeed
set cssText to cssText & ".rc .bar{height:7px;border-radius:6px;background:rgba(255,255,255,.08);overflow:hidden;margin:4px 0 3px}" & linefeed
set cssText to cssText & ".rc .bar i{display:block;height:100%;transition:width .5s}" & linefeed
set cssText to cssText & ".rc .ev{font-size:11px;color:var(--txt-dim);line-height:1.5}" & linefeed
set cssText to cssText & "pre{white-space:pre-wrap;word-break:break-word;background:rgba(0,0,0,.34);border:1px solid var(--border-soft);border-radius:8px;padding:11px;margin:0;font-family:var(--mono);font-size:11px;color:var(--txt-dim);max-height:340px;overflow:auto}" & linefeed
set cssText to cssText & "html[data-theme='light'] pre{background:rgba(0,0,0,.04)}" & linefeed
set cssText to cssText & ".overlay{position:fixed;inset:0;z-index:200;background:rgba(2,5,10,.8);backdrop-filter:blur(7px);display:none;align-items:flex-start;justify-content:center;padding:9vh 16px}" & linefeed
set cssText to cssText & ".overlay.show{display:flex}" & linefeed
set cssText to cssText & ".cmdbox{width:100%;max-width:620px;background:var(--panel-solid);border:1px solid var(--border);border-radius:13px;box-shadow:0 24px 70px rgba(0,0,0,.7);overflow:hidden;animation:pop .2s cubic-bezier(.2,.8,.3,1)}" & linefeed
set cssText to cssText & "@keyframes pop{from{transform:translateY(-14px) scale(.98);opacity:0}to{transform:none;opacity:1}}" & linefeed
set cssText to cssText & ".cmdbox input{width:100%;border:none;background:transparent;padding:15px 17px;font-size:15px;outline:none}" & linefeed
set cssText to cssText & ".cmd-list{max-height:52vh;overflow:auto;border-top:1px solid var(--border-soft)}" & linefeed
set cssText to cssText & ".cmd-item{display:flex;align-items:center;gap:10px;padding:9px 16px;font-size:13px;cursor:pointer;border-left:3px solid transparent}" & linefeed
set cssText to cssText & ".cmd-item .cs{margin-left:auto;font-family:var(--mono);font-size:9px;color:var(--txt-mute);letter-spacing:.1em}" & linefeed
set cssText to cssText & ".cmd-item.sel{background:rgba(0,229,255,.13);border-left-color:var(--cyan)}" & linefeed
set cssText to cssText & ".cmd-foot{display:flex;gap:12px;padding:7px 16px;border-top:1px solid var(--border-soft);font-family:var(--mono);font-size:9px;color:var(--txt-mute);letter-spacing:.08em}" & linefeed
set cssText to cssText & ".modal{width:100%;max-width:660px;background:var(--panel-solid);border:1px solid var(--border);border-radius:13px;box-shadow:0 24px 70px rgba(0,0,0,.7);overflow:hidden;animation:pop .2s;max-height:82vh;display:flex;flex-direction:column}" & linefeed
set cssText to cssText & ".modal h2{margin:0;padding:13px 17px;font-family:var(--mono);font-size:12px;letter-spacing:.16em;border-bottom:1px solid var(--border-soft);display:flex;align-items:center;gap:8px}" & linefeed
set cssText to cssText & ".modal h2 .x{margin-left:auto;background:none;border:none;color:var(--txt-mute);font-size:16px}" & linefeed
set cssText to cssText & ".modal-body{padding:15px 17px;overflow:auto}" & linefeed
set cssText to cssText & "#toasts{position:fixed;bottom:16px;right:16px;z-index:300;display:flex;flex-direction:column;gap:8px;max-width:340px}" & linefeed
set cssText to cssText & ".toast{border:1px solid var(--border);background:var(--panel-solid);border-radius:9px;padding:10px 13px;font-family:var(--mono);font-size:11.5px;box-shadow:var(--shadow);border-left:3px solid var(--cyan);animation:slideIn .25s}" & linefeed
set cssText to cssText & "@keyframes slideIn{from{transform:translateX(40px);opacity:0}to{transform:none;opacity:1}}" & linefeed
set cssText to cssText & ".toast.ok{border-left-color:var(--green)}" & linefeed
set cssText to cssText & ".toast.warn{border-left-color:var(--amber)}" & linefeed
set cssText to cssText & ".toast.err{border-left-color:var(--red)}" & linefeed
set cssText to cssText & "footer{max-width:1680px;margin:20px auto 0;padding:22px 14px 34px;font-family:var(--mono);font-size:10px;color:var(--txt-mute);line-height:1.75;border-top:1px solid var(--border-soft);position:relative;z-index:1}" & linefeed
set cssText to cssText & "@media(max-width:820px){.grid{grid-template-columns:1fr}.panel.wide{grid-column:span 1}.hero-rate{text-align:left}.hero-rate .big{font-size:32px}.hero-id h1{font-size:20px}.wrap{padding:10px}}" & linefeed
set cssText to cssText & "@media print{body{background:#fff;color:#000}.topbar,.subbar,.overlay,#toasts,.p-act,.noprint{display:none!important}.panel{break-inside:avoid;border:1px solid #999}.p-body{max-height:none!important}}"

-- ------------------------------------------------------------ page shell ---

set shellText to ""
set shellText to shellText & "</style></head><body>" & linefeed
set shellText to shellText & "<div class='topbar' id='topbar'>" & linefeed
set shellText to shellText & "<div class='brand'><span class='bolt'>&#128225;</span><span class='b1'>WIFI</span><span class='b2'>CONNECTIVITY MONITOR</span><span class='ver'>v3.0</span></div>" & linefeed
set shellText to shellText & "<span class='live'><span class='dot'></span>CAPTURED</span>" & linefeed
set shellText to shellText & "<span class='statuspill sp-ok' id='locPill'>LOCATION</span>" & linefeed
set shellText to shellText & "<span class='statuspill sp-ok' id='linkPill'>LINK</span>" & linefeed
set shellText to shellText & "<div class='tb-spacer'></div>" & linefeed
set shellText to shellText & "<button class='tb-btn searchbtn' data-act='palette'>&#128269; Search / Commands <span class='kbd'>&#8984;K</span></button>" & linefeed
set shellText to shellText & "<button class='tb-btn' data-act='theme' title='Toggle theme (T)'>&#9680;</button>" & linefeed
set shellText to shellText & "<button class='tb-btn' data-act='rerender' title='Re-render panels (R)'>&#8635;</button>" & linefeed
set shellText to shellText & "<button class='tb-btn' data-act='export' title='Export snapshot (E)'>&#10515; EXPORT</button>" & linefeed
set shellText to shellText & "<button class='tb-btn' data-act='print' title='Print / PDF (P)'>&#128424; PDF</button>" & linefeed
set shellText to shellText & "<button class='tb-btn' data-act='full' title='Fullscreen (F)'>&#9974;</button>" & linefeed
set shellText to shellText & "<div class='clockbox'><b id='clockLocal'>--:--:--</b><span id='clockMeta'>CAPTURE</span></div>" & linefeed
set shellText to shellText & "</div>" & linefeed
set shellText to shellText & "<div class='subbar' id='subbar'>" & linefeed
set shellText to shellText & "<span class='sb-label'>PANELS</span>" & linefeed
set shellText to shellText & "<div id='chipRow' style='display:flex;gap:6px'></div>" & linefeed
set shellText to shellText & "<div class='range-pills' id='rangePills'>" & linefeed
set shellText to shellText & "<span class='sb-label'>SAMPLES</span>" & linefeed
set shellText to shellText & "<button class='pill' data-win='25'>25</button>" & linefeed
set shellText to shellText & "<button class='pill' data-win='50'>50</button>" & linefeed
set shellText to shellText & "<button class='pill' data-win='100'>100</button>" & linefeed
set shellText to shellText & "<button class='pill on' data-win='0'>ALL</button>" & linefeed
set shellText to shellText & "</div>" & linefeed
set shellText to shellText & "</div>" & linefeed
set shellText to shellText & "<div class='wrap'>" & linefeed
set shellText to shellText & "<div id='alertBanner'><span style='font-size:16px'>&#9888;</span><span id='alertText'></span><button class='x' data-act='dismiss'>&#10005;</button></div>" & linefeed
set shellText to shellText & "<div class='hero'>" & linefeed
set shellText to shellText & "<div class='hero-top'>" & linefeed
set shellText to shellText & "<div class='statuslamp lamp-ok' id='lamp'>&#128994;</div>" & linefeed
set shellText to shellText & "<div class='hero-id'>" & linefeed
set shellText to shellText & "<div class='lbl'>CAPTURE VERDICT</div>" & linefeed
set shellText to shellText & "<h1 id='heroTitle'>ANALYZING</h1>" & linefeed
set shellText to shellText & "<div class='sub' id='heroSub'></div>" & linefeed
set shellText to shellText & "</div>" & linefeed
set shellText to shellText & "<div class='hero-rate'>" & linefeed
set shellText to shellText & "<div class='big' id='heroBig'>0%</div>" & linefeed
set shellText to shellText & "<div class='unit'>PACKET LOSS TO 8.8.8.8</div>" & linefeed
set shellText to shellText & "<div style='margin-top:6px'><span class='statuspill sp-ok' id='heroPill'>--</span></div>" & linefeed
set shellText to shellText & "</div>" & linefeed
set shellText to shellText & "</div>" & linefeed
set shellText to shellText & "<div class='kpis' id='kpis'></div>" & linefeed
set shellText to shellText & "</div>" & linefeed
set shellText to shellText & "<div class='grid' id='grid'></div>" & linefeed
set shellText to shellText & "</div>" & linefeed
set shellText to shellText & "<footer>" & linefeed
set shellText to shellText & "<div><b class='c-cyan'>WIFI CONNECTIVITY MONITOR v3.0</b> &middot; IT Support &middot; read-only capture, no network settings were changed.</div>" & linefeed
set shellText to shellText & "<div>Root cause scores are weighted heuristics derived from the measurements in this capture. They rank where to investigate first; they are not a substitute for switch, controller or ISP-side telemetry.</div>" & linefeed
set shellText to shellText & "<div>Signal, channel and neighbour data are read from system_profiler SPAirPortDataType. If SSID reads as unavailable, grant this app Location Services access &mdash; all other metrics are unaffected.</div>" & linefeed
set shellText to shellText & "<div class='noprint'>100% local &middot; nothing leaves this Mac &middot; capture ledger stored in this browser profile localStorage &middot; <span class='kbd'>&#8984;K</span> commands &middot; <span class='kbd'>?</span> shortcuts.</div>" & linefeed
set shellText to shellText & "</footer>" & linefeed
set shellText to shellText & "<div class='overlay' id='palette'><div class='cmdbox'><input id='cmdInput' placeholder='Jump to a panel or run a command...' autocomplete='off' spellcheck='false'><div class='cmd-list' id='cmdList'></div><div class='cmd-foot'><span>UP/DOWN NAVIGATE</span><span>ENTER RUN</span><span>ESC CLOSE</span></div></div></div>" & linefeed
set shellText to shellText & "<div class='overlay' id='helpModal'><div class='modal'><h2>&#9000; KEYBOARD SHORTCUTS <button class='x' data-act='closeHelp'>&#10005;</button></h2><div class='modal-body'><table><tbody>" & linefeed
set shellText to shellText & "<tr><td><span class='kbd'>&#8984;K</span> / <span class='kbd'>Ctrl K</span></td><td class='dim'>Open command palette</td></tr>" & linefeed
set shellText to shellText & "<tr><td><span class='kbd'>T</span></td><td class='dim'>Toggle dark / light theme</td></tr>" & linefeed
set shellText to shellText & "<tr><td><span class='kbd'>F</span></td><td class='dim'>Toggle fullscreen</td></tr>" & linefeed
set shellText to shellText & "<tr><td><span class='kbd'>R</span></td><td class='dim'>Re-render all panels</td></tr>" & linefeed
set shellText to shellText & "<tr><td><span class='kbd'>E</span></td><td class='dim'>Export capture snapshot (JSON)</td></tr>" & linefeed
set shellText to shellText & "<tr><td><span class='kbd'>P</span></td><td class='dim'>Print / save as PDF</td></tr>" & linefeed
set shellText to shellText & "<tr><td><span class='kbd'>C</span></td><td class='dim'>Copy the Copilot root cause prompt</td></tr>" & linefeed
set shellText to shellText & "<tr><td><span class='kbd'>1..4</span></td><td class='dim'>Sample window (25 / 50 / 100 / ALL)</td></tr>" & linefeed
set shellText to shellText & "<tr><td><span class='kbd'>?</span></td><td class='dim'>This help</td></tr>" & linefeed
set shellText to shellText & "<tr><td><span class='kbd'>ESC</span></td><td class='dim'>Close any overlay</td></tr>" & linefeed
set shellText to shellText & "</tbody></table></div></div></div>" & linefeed
set shellText to shellText & "<div id='toasts'></div>" & linefeed
set shellText to shellText & "<script>"

-- ------------------------------------------------------ dashboard engine ---

set jsText to ""
set jsText to jsText & "'use strict';" & linefeed
set jsText to jsText & "var LS={get:function(k,d){try{var v=localStorage.getItem('wifimon3.'+k);return v===null?d:JSON.parse(v);}catch(e){return d;}},set:function(k,v){try{localStorage.setItem('wifimon3.'+k,JSON.stringify(v));}catch(e){}}};" & linefeed
set jsText to jsText & "function esc(s){if(s===null||s===undefined)return '';s=String(s);return s.split('&').join('&amp;').split('<').join('&lt;').split('>').join('&gt;').split(String.fromCharCode(34)).join('&quot;').split(String.fromCharCode(39)).join('&#39;');}" & linefeed
set jsText to jsText & "function q(s){return document.querySelector(s);}" & linefeed
set jsText to jsText & "function chanNum(t){var m=String(t).match(/[0-9]+/);return m?parseInt(m[0],10):0;}" & linefeed
set jsText to jsText & "function bandOf(c){if(c<=0)return 'unknown';if(c<=14)return '2.4 GHz';if(c<=177)return '5 GHz';return '6 GHz';}" & linefeed
set jsText to jsText & "function fx(v,n){return (Math.round(v*Math.pow(10,n))/Math.pow(10,n)).toFixed(n);}" & linefeed
set jsText to jsText & "var WIN=LS.get('win',0);" & linefeed
set jsText to jsText & "var CHAN=chanNum(CAP.chan),BAND=bandOf(CHAN),ISWIFI=(CAP.link.indexOf('Wi-Fi')===0);" & linefeed
set jsText to jsText & "var Ledger={all:function(){return LS.get('caps',[]);}," & linefeed
set jsText to jsText & "save:function(){var a=this.all();a=a.filter(function(x){return x.id!==CAP.id;});a.push({id:CAP.id,ts:CAP.ts,loc:CAP.loc,link:CAP.link,ssid:CAP.ssid,chan:CAP.chan,rssi:CAP.rssi,snr:CAP.snr,haveSig:CAP.haveSig,loss:CAP.loss,avg:CAP.avg,max:CAP.max,jit:CAP.jit,drop:CAP.drop,nbTotal:CAP.nbTotal,gw:CAP.gw,n:CAP.n});a.sort(function(x,y){return x.ts-y.ts;});if(a.length>14)a=a.slice(a.length-14);LS.set('caps',a);return a;}," & linefeed
set jsText to jsText & "clear:function(){LS.set('caps',[]);}};" & linefeed
set jsText to jsText & "var CAPS=Ledger.save();" & linefeed
set jsText to jsText & "function toast(m,kind,ms){var t=document.createElement('div');t.className='toast '+(kind||'');t.textContent=m;q('#toasts').appendChild(t);setTimeout(function(){t.style.transition='opacity .3s,transform .3s';t.style.opacity='0';t.style.transform='translateX(40px)';setTimeout(function(){t.remove();},320);},ms||4200);}" & linefeed
set jsText to jsText & "var Score={run:function(){" & linefeed
set jsText to jsText & "var s=[],ev,sc;" & linefeed
set jsText to jsText & "sc=0;ev=[];" & linefeed
set jsText to jsText & "if(CAP.haveSig&&CAP.snr>0&&CAP.snr<25){sc+=45;ev.push('SNR is '+CAP.snr+' dB. Healthy is 25 dB or better, so the noise floor here is elevated.');}" & linefeed
set jsText to jsText & "else if(CAP.haveSig&&CAP.snr>=25&&CAP.snr<32){sc+=18;ev.push('SNR '+CAP.snr+' dB is workable but not comfortable.');}" & linefeed
set jsText to jsText & "if(CAP.nbTotal>12){sc+=25;ev.push(CAP.nbTotal+' neighbouring networks detected. This is a dense multi-tenant RF environment.');}" & linefeed
set jsText to jsText & "else if(CAP.nbTotal>6){sc+=12;ev.push(CAP.nbTotal+' neighbouring networks detected.');}" & linefeed
set jsText to jsText & "if(BAND==='2.4 GHz'){sc+=18;ev.push('Client is associated on 2.4 GHz (channel '+CHAN+'), the most congested and interference-prone band.');}" & linefeed
set jsText to jsText & "if(ISWIFI&&CAP.jit>25){sc+=12;ev.push('Jitter of '+CAP.jit+' ms suggests contention for airtime.');}" & linefeed
set jsText to jsText & "if(!ISWIFI){sc=Math.round(sc*0.15);ev=['This capture ran over a wired link, so wireless interference is largely excluded here.'];}" & linefeed
set jsText to jsText & "s.push({k:'RF Interference / Noise',sc:sc,ev:ev});" & linefeed
set jsText to jsText & "sc=0;ev=[];" & linefeed
set jsText to jsText & "if(CAP.haveSig&&CAP.rssi<=-78){sc+=55;ev.push('RSSI '+CAP.rssi+' dBm is very weak. The client is at the edge of usable coverage.');}" & linefeed
set jsText to jsText & "else if(CAP.haveSig&&CAP.rssi<=-72){sc+=35;ev.push('RSSI '+CAP.rssi+' dBm is marginal for this position.');}" & linefeed
set jsText to jsText & "else if(CAP.haveSig&&CAP.rssi<=-67){sc+=15;ev.push('RSSI '+CAP.rssi+' dBm is acceptable but not strong.');}" & linefeed
set jsText to jsText & "if(CAP.haveSig&&CAP.rssi<=-72&&CAP.loss>0){sc+=20;ev.push('Weak signal is coinciding with measured packet loss.');}" & linefeed
set jsText to jsText & "if(!ISWIFI){sc=0;ev=['Wired capture. Access point coverage is not a factor for this measurement.'];}" & linefeed
set jsText to jsText & "s.push({k:'Access Point Coverage / Placement',sc:sc,ev:ev});" & linefeed
set jsText to jsText & "sc=0;ev=[];" & linefeed
set jsText to jsText & "if(CAP.drop>=5){sc+=60;ev.push('Longest unbroken outage was '+CAP.drop+' consecutive samples. That is a genuine link drop, not latency.');}" & linefeed
set jsText to jsText & "else if(CAP.drop>=3){sc+=40;ev.push(CAP.drop+' consecutive samples failed, consistent with a roaming event or association loss.');}" & linefeed
set jsText to jsText & "else if(CAP.drop===2){sc+=18;ev.push('Two consecutive samples failed, so a brief interruption was observed.');}" & linefeed
set jsText to jsText & "if(ISWIFI&&CAP.drop>=3&&CAP.haveSig&&CAP.rssi>-70){sc+=15;ev.push('Drops are occurring despite adequate signal. Look at roaming thresholds and access point overlap.');}" & linefeed
set jsText to jsText & "s.push({k:'Roaming / Association Loss',sc:sc,ev:ev});" & linefeed
set jsText to jsText & "sc=0;ev=[];" & linefeed
set jsText to jsText & "if(CAP.gw==='unreachable'){sc+=70;ev.push('The default gateway did not respond at all.');}" & linefeed
set jsText to jsText & "else if(CAP.gw==='lossy'){sc+=45;ev.push('The default gateway responded but with packet loss, so the problem starts inside the office.');}" & linefeed
set jsText to jsText & "else if(CAP.gw==='unknown'){sc+=25;ev.push('No default gateway was detected on this interface.');}" & linefeed
set jsText to jsText & "if(CAP.hop1===1){sc+=20;ev.push('The first traceroute hop did not answer.');}" & linefeed
set jsText to jsText & "if(CAP.gw==='clean'&&CAP.loss>0){ev.push('The gateway itself is clean, so the local router is unlikely to be the origin.');}" & linefeed
set jsText to jsText & "s.push({k:'Local Router / Gateway',sc:sc,ev:ev});" & linefeed
set jsText to jsText & "sc=0;ev=[];" & linefeed
set jsText to jsText & "if(CAP.dnsOk===0){sc+=65;ev.push('DNS resolution for the test domain did not return an answer.');}" & linefeed
set jsText to jsText & "if(CAP.httpCode==='000'){sc+=35;ev.push('The HTTPS reachability test did not complete. Name resolution or egress is blocked or timing out.');}" & linefeed
set jsText to jsText & "else if(CAP.httpCode!=='200'){sc+=20;ev.push('HTTPS reachability test returned status '+CAP.httpCode+'.');}" & linefeed
set jsText to jsText & "if(CAP.dns==='Not detected'){sc+=20;ev.push('No DNS servers were configured on the active interface.');}" & linefeed
set jsText to jsText & "s.push({k:'DNS / Name Resolution',sc:sc,ev:ev});" & linefeed
set jsText to jsText & "sc=0;ev=[];" & linefeed
set jsText to jsText & "if(!ISWIFI&&CAP.loss>0){sc+=65;ev.push('Loss is present on a WIRED connection. This points upstream of the wireless network entirely.');}" & linefeed
set jsText to jsText & "if(CAP.gw==='clean'&&CAP.loss>=2){sc+=40;ev.push('The gateway is clean but internet-bound traffic is losing '+CAP.loss+' percent, so the break is beyond the office router.');}" & linefeed
set jsText to jsText & "if(CAP.max>250&&CAP.avg<120){sc+=15;ev.push('Latency spiked to '+CAP.max+' ms against a '+CAP.avg+' ms average, which is typical of upstream congestion.');}" & linefeed
set jsText to jsText & "if(CAP.avg>90){sc+=15;ev.push('Average latency of '+CAP.avg+' ms is high for a business circuit.');}" & linefeed
set jsText to jsText & "s.push({k:'ISP / Upstream Circuit',sc:sc,ev:ev});" & linefeed
set jsText to jsText & "sc=0;ev=[];" & linefeed
set jsText to jsText & "if(CAP.ssid==='Not available'&&ISWIFI){sc+=10;ev.push('SSID could not be read. This is usually a Location Services permission, not a fault.');}" & linefeed
set jsText to jsText & "if(CAP.n<40){sc+=8;ev.push('Only '+CAP.n+' samples were taken. Run an Extended or Deep capture before ruling anything out.');}" & linefeed
set jsText to jsText & "if(CAP.loss===0&&CAP.drop===0){ev.push('No endpoint-level packet loss was observed during this capture window.');}" & linefeed
set jsText to jsText & "s.push({k:'Endpoint / Client Device',sc:sc,ev:ev});" & linefeed
set jsText to jsText & "s.forEach(function(x){if(x.sc>100)x.sc=100;});" & linefeed
set jsText to jsText & "s.sort(function(a,b){return b.sc-a.sc;});" & linefeed
set jsText to jsText & "return s;}};" & linefeed
set jsText to jsText & "var SCORES=Score.run();" & linefeed
set jsText to jsText & "var VERDICT=(function(){var top=SCORES[0];" & linefeed
set jsText to jsText & "if(CAP.loss===0&&CAP.drop===0&&(!CAP.haveSig||CAP.rssi>-70))return {t:'LINK HEALTHY',p:'NO FAULT OBSERVED',lamp:'lamp-ok',pill:'sp-ok',emo:String.fromCodePoint(128994)};" & linefeed
set jsText to jsText & "if(CAP.loss>=5||CAP.drop>=3)return {t:'LINK DEGRADED',p:top.sc>0?('LIKELY: '+top.k.toUpperCase()):'INVESTIGATE',lamp:'lamp-bad',pill:'sp-bad',emo:String.fromCodePoint(128308)};" & linefeed
set jsText to jsText & "if(CAP.loss>0||(CAP.haveSig&&CAP.rssi<=-72)||CAP.jit>30)return {t:'MARGINAL LINK',p:top.sc>0?('WATCH: '+top.k.toUpperCase()):'MONITOR',lamp:'lamp-warn',pill:'sp-warn',emo:String.fromCodePoint(128993)};" & linefeed
set jsText to jsText & "return {t:'LINK STABLE',p:'WITHIN TOLERANCE',lamp:'lamp-ok',pill:'sp-ok',emo:String.fromCodePoint(128994)};})();" & linefeed
set jsText to jsText & "var Chart={" & linefeed
set jsText to jsText & "prep:function(cv,h){var dpr=window.devicePixelRatio||1,w=cv.clientWidth||360;cv.width=w*dpr;cv.height=h*dpr;cv.style.height=h+'px';var c=cv.getContext('2d');c.setTransform(dpr,0,0,dpr,0,0);c.clearRect(0,0,w,h);return {c:c,w:w,h:h};}," & linefeed
set jsText to jsText & "css:function(v){return getComputedStyle(document.documentElement).getPropertyValue(v).trim();}," & linefeed
set jsText to jsText & "latency:function(cv){var o=this.prep(cv,180),c=o.c,w=o.w,h=o.h;var d=CAP.samples;if(WIN>0&&d.length>WIN)d=d.slice(d.length-WIN);" & linefeed
set jsText to jsText & "if(!d.length){c.fillStyle=this.css('--txt-mute');c.font='11px '+this.css('--mono');c.textAlign='center';c.fillText('No samples',w/2,h/2);return;}" & linefeed
set jsText to jsText & "var ok=d.filter(function(x){return x.ok;});var mx=ok.length?Math.max.apply(null,ok.map(function(x){return x.ms;})):1;mx=Math.max(mx*1.15,10);" & linefeed
set jsText to jsText & "var pad=34,plotH=h-40;" & linefeed
set jsText to jsText & "c.strokeStyle='rgba(125,147,168,.13)';c.lineWidth=1;" & linefeed
set jsText to jsText & "for(var g=0;g<=4;g++){var y=10+plotH*(g/4);c.beginPath();c.moveTo(pad,y);c.lineTo(w-4,y);c.stroke();c.fillStyle=this.css('--txt-mute');c.font='8px '+this.css('--mono');c.textAlign='right';c.fillText(Math.round(mx*(1-g/4))+'',pad-4,y+3);}" & linefeed
set jsText to jsText & "var step=(w-pad-6)/Math.max(1,d.length-1);" & linefeed
set jsText to jsText & "d.forEach(function(s,i){if(s.ok)return;var x=pad+i*step;c.fillStyle='rgba(255,59,87,.22)';c.fillRect(x-Math.max(1.5,step/2),10,Math.max(3,step),plotH);});" & linefeed
set jsText to jsText & "c.beginPath();var started=false;" & linefeed
set jsText to jsText & "d.forEach(function(s,i){var x=pad+i*step;if(!s.ok){started=false;return;}var y=10+plotH-(s.ms/mx)*plotH;if(!started){c.moveTo(x,y);started=true;}else{c.lineTo(x,y);}});" & linefeed
set jsText to jsText & "c.strokeStyle=this.css('--cyan');c.lineWidth=1.8;c.stroke();" & linefeed
set jsText to jsText & "var avgY=10+plotH-(CAP.avg/mx)*plotH;c.strokeStyle=this.css('--violet');c.setLineDash([4,3]);c.lineWidth=1.1;c.beginPath();c.moveTo(pad,avgY);c.lineTo(w-4,avgY);c.stroke();c.setLineDash([]);" & linefeed
set jsText to jsText & "c.fillStyle=this.css('--violet');c.font='700 8px '+this.css('--mono');c.textAlign='left';c.fillText('AVG '+CAP.avg+' ms',pad+3,avgY-4);" & linefeed
set jsText to jsText & "c.fillStyle=this.css('--txt-mute');c.font='8px '+this.css('--mono');c.textAlign='left';c.fillText(d[0].t,pad,h-4);c.textAlign='right';c.fillText(d[d.length-1].t,w-4,h-4);c.textAlign='center';c.fillText(d.length+' samples   red bands are failed pings',w/2,h-4);}," & linefeed
set jsText to jsText & "signal:function(cv){var o=this.prep(cv,150),c=o.c,w=o.w,h=o.h,self=this;" & linefeed
set jsText to jsText & "if(!CAP.haveSig){c.fillStyle=this.css('--txt-mute');c.font='11px '+this.css('--mono');c.textAlign='center';c.fillText('Signal metrics unavailable on this link',w/2,h/2);return;}" & linefeed
set jsText to jsText & "var rows=[['RSSI',CAP.rssi,-95,-35,[-72,-67]],['NOISE',CAP.noise,-100,-40,[-90,-80]],['SNR',CAP.snr,0,50,[20,25]]];" & linefeed
set jsText to jsText & "var barH=18,gap=28,x0=54,bw=w-x0-64;" & linefeed
set jsText to jsText & "rows.forEach(function(r,i){var y=20+i*gap;var t=(r[1]-r[2])/(r[3]-r[2]);if(t<0)t=0;if(t>1)t=1;" & linefeed
set jsText to jsText & "c.fillStyle=self.css('--txt-mute');c.font='9px '+self.css('--mono');c.textAlign='left';c.fillText(r[0],4,y+13);" & linefeed
set jsText to jsText & "c.fillStyle='rgba(255,255,255,.07)';c.fillRect(x0,y,bw,barH);" & linefeed
set jsText to jsText & "var col=self.css('--green');" & linefeed
set jsText to jsText & "if(r[0]==='SNR'){if(r[1]<r[4][0])col=self.css('--red');else if(r[1]<r[4][1])col=self.css('--amber');}" & linefeed
set jsText to jsText & "else if(r[0]==='RSSI'){if(r[1]<=r[4][0])col=self.css('--red');else if(r[1]<=r[4][1])col=self.css('--amber');}" & linefeed
set jsText to jsText & "else{if(r[1]>r[4][1])col=self.css('--red');else if(r[1]>r[4][0])col=self.css('--amber');}" & linefeed
set jsText to jsText & "c.fillStyle=col;c.fillRect(x0,y,bw*t,barH);" & linefeed
set jsText to jsText & "c.fillStyle=self.css('--txt');c.font='700 11px '+self.css('--mono');c.textAlign='right';c.fillText(r[1]+(r[0]==='SNR'?' dB':' dBm'),w-4,y+13);});" & linefeed
set jsText to jsText & "c.fillStyle=this.css('--txt-mute');c.font='8px '+this.css('--mono');c.textAlign='left';c.fillText('targets: RSSI better than -67 dBm, noise below -90 dBm, SNR above 25 dB',4,h-8);}," & linefeed
set jsText to jsText & "chan:function(cv){var o=this.prep(cv,150),c=o.c,w=o.w,h=o.h,self=this;var d=CAP.nb.slice(0,14);" & linefeed
set jsText to jsText & "if(!d.length){c.fillStyle=this.css('--txt-mute');c.font='11px '+this.css('--mono');c.textAlign='center';c.fillText('No neighbouring networks reported',w/2,h/2);return;}" & linefeed
set jsText to jsText & "d.sort(function(a,b){return a.ch-b.ch;});" & linefeed
set jsText to jsText & "var mx=Math.max.apply(null,d.map(function(x){return x.n;}));var plotH=h-36,bw=(w-14)/d.length;" & linefeed
set jsText to jsText & "d.forEach(function(x,i){var bh=(x.n/mx)*plotH,px=7+i*bw,py=12+plotH-bh;var mine=(x.ch===CHAN);" & linefeed
set jsText to jsText & "var col=mine?self.css('--cyan'):(x.ch<=14?self.css('--amber'):self.css('--violet'));" & linefeed
set jsText to jsText & "var g=c.createLinearGradient(0,py,0,12+plotH);g.addColorStop(0,col);g.addColorStop(1,'rgba(125,147,168,.08)');" & linefeed
set jsText to jsText & "c.fillStyle=g;c.fillRect(px+1.5,py,bw-3,Math.max(2,bh));" & linefeed
set jsText to jsText & "if(mine){c.strokeStyle='#fff';c.lineWidth=1.4;c.strokeRect(px+1.5,py,bw-3,Math.max(2,bh));}" & linefeed
set jsText to jsText & "c.fillStyle=self.css('--txt');c.font='700 9px '+self.css('--mono');c.textAlign='center';c.fillText(x.n+'',px+bw/2,py-4);" & linefeed
set jsText to jsText & "c.fillStyle=mine?self.css('--cyan'):self.css('--txt-mute');c.font='8px '+self.css('--mono');c.fillText(x.ch+'',px+bw/2,h-14);});" & linefeed
set jsText to jsText & "c.fillStyle=this.css('--txt-mute');c.font='8px '+this.css('--mono');c.textAlign='center';c.fillText('networks per channel   cyan is your channel, amber is 2.4 GHz, violet is 5 or 6 GHz',w/2,h-3);}," & linefeed
set jsText to jsText & "compare:function(cv){var o=this.prep(cv,160),c=o.c,w=o.w,h=o.h,self=this;var d=CAPS.slice(-8);" & linefeed
set jsText to jsText & "if(d.length<2){c.fillStyle=this.css('--txt-mute');c.font='11px '+this.css('--mono');c.textAlign='center';c.fillText('Run a second capture at another position to compare',w/2,h/2);return;}" & linefeed
set jsText to jsText & "var mx=Math.max(1,Math.max.apply(null,d.map(function(x){return x.loss;})));var plotH=h-44,bw=(w-14)/d.length;" & linefeed
set jsText to jsText & "d.forEach(function(x,i){var bh=(x.loss/mx)*plotH,px=7+i*bw,py=14+plotH-bh;" & linefeed
set jsText to jsText & "var col=x.loss>=5?self.css('--red'):(x.loss>0?self.css('--amber'):self.css('--green'));" & linefeed
set jsText to jsText & "c.fillStyle=col;c.fillRect(px+2,py,bw-4,Math.max(2,bh));" & linefeed
set jsText to jsText & "if(x.id===CAP.id){c.strokeStyle='#fff';c.lineWidth=1.4;c.strokeRect(px+2,py,bw-4,Math.max(2,bh));}" & linefeed
set jsText to jsText & "c.fillStyle=self.css('--txt');c.font='700 9px '+self.css('--mono');c.textAlign='center';c.fillText(fx(x.loss,1)+'%',px+bw/2,py-4);" & linefeed
set jsText to jsText & "c.fillStyle=self.css('--txt-mute');c.font='7.5px '+self.css('--mono');" & linefeed
set jsText to jsText & "c.fillText(x.loc.split(' ')[0].substring(0,10),px+bw/2,h-14);" & linefeed
set jsText to jsText & "c.fillText(x.link.indexOf('Wi-Fi')===0?'wifi':'wired',px+bw/2,h-4);});" & linefeed
set jsText to jsText & "c.fillStyle=this.css('--txt-mute');c.font='8px '+this.css('--mono');c.textAlign='left';c.fillText('packet loss by capture point',4,10);}};" & linefeed
set jsText to jsText & "var Panels={hidden:LS.get('hidden',{}),collapsed:LS.get('collapsed',{})," & linefeed
set jsText to jsText & "defs:[{id:'verdict',ico:String.fromCodePoint(127919),title:'ROOT CAUSE MATRIX',wide:true,r:'rVerdict'}," & linefeed
set jsText to jsText & "{id:'latency',ico:String.fromCodePoint(128200),title:'LATENCY AND DROP TIMELINE',wide:true,r:'rLatency'}," & linefeed
set jsText to jsText & "{id:'signal',ico:String.fromCodePoint(128246),title:'SIGNAL QUALITY',wide:false,r:'rSignal'}," & linefeed
set jsText to jsText & "{id:'chan',ico:String.fromCodePoint(128225),title:'CHANNEL CONGESTION',wide:false,r:'rChan'}," & linefeed
set jsText to jsText & "{id:'compare',ico:String.fromCodePoint(9878),title:'CAPTURE COMPARISON',wide:true,r:'rCompare'}," & linefeed
set jsText to jsText & "{id:'path',ico:String.fromCodePoint(128279),title:'PATH VALIDATION',wide:false,r:'rPath'}," & linefeed
set jsText to jsText & "{id:'actions',ico:String.fromCodePoint(9989),title:'RECOMMENDED ACTIONS',wide:false,r:'rActions'}," & linefeed
set jsText to jsText & "{id:'device',ico:String.fromCodePoint(128187),title:'DEVICE AND INTERFACE',wide:false,r:'rDevice'}," & linefeed
set jsText to jsText & "{id:'samples',ico:String.fromCodePoint(128203),title:'SAMPLE LOG',wide:false,r:'rSamples'}," & linefeed
set jsText to jsText & "{id:'copilot',ico:String.fromCodePoint(129302),title:'COPILOT PROMPTS',wide:true,r:'rCopilot'}," & linefeed
set jsText to jsText & "{id:'raw',ico:String.fromCodePoint(128220),title:'RAW DIAGNOSTIC OUTPUT',wide:true,r:'rRaw'}]," & linefeed
set jsText to jsText & "visible:function(){var self=this;return this.defs.filter(function(d){return !self.hidden[d.id];});}," & linefeed
set jsText to jsText & "toggle:function(id){this.hidden[id]=!this.hidden[id];LS.set('hidden',this.hidden);this.build();this.chips();}," & linefeed
set jsText to jsText & "collapse:function(id){this.collapsed[id]=!this.collapsed[id];LS.set('collapsed',this.collapsed);var el=document.getElementById('panel_'+id);if(el)el.classList.toggle('collapsed',!!this.collapsed[id]);}," & linefeed
set jsText to jsText & "focus:function(id){if(this.hidden[id])this.toggle(id);var el=document.getElementById('panel_'+id);if(!el)return;if(this.collapsed[id])this.collapse(id);el.scrollIntoView({behavior:'smooth',block:'center'});el.classList.add('glow');setTimeout(function(){el.classList.remove('glow');},1800);}," & linefeed
set jsText to jsText & "chips:function(){var self=this;q('#chipRow').innerHTML=this.defs.map(function(d){return '<button class=\"chip '+(self.hidden[d.id]?'':'on')+'\" data-act=\"chip\" data-id=\"'+d.id+'\">'+d.ico+' '+d.title.split(' ')[0]+'</button>';}).join('');}," & linefeed
set jsText to jsText & "build:function(){var self=this;q('#grid').innerHTML=this.visible().map(function(d){return '<section class=\"panel '+(d.wide?'wide':'')+(self.collapsed[d.id]?' collapsed':'')+'\" id=\"panel_'+d.id+'\"><div class=\"p-head\" data-act=\"collapse\" data-id=\"'+d.id+'\"><h3><span class=\"caret\">&#9662;</span><span>'+d.ico+'</span>'+d.title+'</h3><button class=\"p-act\" data-act=\"hide\" data-id=\"'+d.id+'\">&#10005;</button></div><div class=\"p-body\" id=\"body_'+d.id+'\"></div></section>';}).join('');this.renderAll();}," & linefeed
set jsText to jsText & "render:function(id){var d=null;this.defs.forEach(function(x){if(x.id===id)d=x;});if(!d)return;var b=document.getElementById('body_'+id);if(!b)return;b.innerHTML=Render[d.r]();if(Render[d.r+'_after'])Render[d.r+'_after']();}," & linefeed
set jsText to jsText & "renderAll:function(){var self=this;this.visible().forEach(function(d){self.render(d.id);});}};" & linefeed
set jsText to jsText & "var Render={" & linefeed
set jsText to jsText & "rVerdict:function(){var h='<div class=\"hint\" style=\"margin-bottom:10px\">Weighted ranking of where to investigate first, computed from this capture only. Bars show relative confidence, not probability.</div>';" & linefeed
set jsText to jsText & "SCORES.forEach(function(s){var col=s.sc>=50?'--red':(s.sc>=25?'--amber':(s.sc>0?'--violet':'--green'));" & linefeed
set jsText to jsText & "h+='<div class=\"rc\"><div class=\"rc-top\"><b>'+esc(s.k)+'</b><span class=\"mono\" style=\"color:var('+col+')\">'+s.sc+'</span></div><div class=\"bar\"><i style=\"width:'+Math.max(2,s.sc)+'%;background:var('+col+')\"></i></div><div class=\"ev\">'+(s.ev.length?s.ev.map(function(e){return '&bull; '+esc(e);}).join('<br>'):'&bull; No supporting evidence in this capture.')+'</div></div>';});" & linefeed
set jsText to jsText & "var top=SCORES[0];var cls=top.sc>=50?'bad':(top.sc>=25?'warn':'good');" & linefeed
set jsText to jsText & "var msg=top.sc<=0?'Nothing in this capture indicates a fault. If the user still reports drops, repeat with a Deep capture during an active Zoom or Teams call.':('Start with <b>'+esc(top.k)+'</b>. '+esc(top.ev[0]||''));" & linefeed
set jsText to jsText & "return h+'<div class=\"advice '+cls+'\">'+msg+'</div>';}," & linefeed
set jsText to jsText & "rLatency:function(){return '<canvas id=\"cvLat\"></canvas><div class=\"result\"><div class=\"line\"><span>Samples sent</span><b>'+CAP.n+' to '+esc(CAP.target)+'</b></div><div class=\"line\"><span>Replies / timeouts</span><b><span class=\"c-green\">'+CAP.ok+'</span> / <span class=\"c-red\">'+CAP.bad+'</span></b></div><div class=\"line\"><span>Packet loss</span><b class=\"'+(CAP.loss>0?'c-red':'c-green')+'\">'+CAP.loss+'%</b></div><div class=\"line\"><span>Latency min / avg / max</span><b>'+CAP.min+' / '+CAP.avg+' / '+CAP.max+' ms</b></div><div class=\"line\"><span>Jitter (mean deviation)</span><b class=\"'+(CAP.jit>30?'c-amber':'')+'\">'+CAP.jit+' ms</b></div><div class=\"line\"><span>Longest unbroken outage</span><b class=\"'+(CAP.drop>=3?'c-red':'')+'\">'+CAP.drop+' samples</b></div></div><div class=\"hint\" style=\"margin-top:8px\">A single failed sample is usually noise. Three or more in a row is a real link drop, and that is what users describe as the connection cutting out.</div>';}," & linefeed
set jsText to jsText & "rLatency_after:function(){var cv=document.getElementById('cvLat');if(cv)Chart.latency(cv);}," & linefeed
set jsText to jsText & "rSignal:function(){var h='<canvas id=\"cvSig\"></canvas><div class=\"result\"><div class=\"line\"><span>SSID</span><b>'+esc(CAP.ssid)+'</b></div><div class=\"line\"><span>BSSID (radio)</span><b>'+esc(CAP.bssid)+'</b></div><div class=\"line\"><span>Channel / band</span><b>'+esc(CAP.chan)+' &middot; '+BAND+'</b></div><div class=\"line\"><span>PHY mode</span><b>'+esc(CAP.phy)+'</b></div><div class=\"line\"><span>Security</span><b>'+esc(CAP.sec)+'</b></div><div class=\"line\"><span>Transmit rate</span><b>'+esc(CAP.tx)+'</b></div></div>';" & linefeed
set jsText to jsText & "if(!ISWIFI)h+='<div class=\"advice good\">This capture ran over <b>'+esc(CAP.link)+'</b>. Comparing it against a Wi-Fi capture from the same desk is the fastest way to separate wireless problems from router or ISP problems.</div>';" & linefeed
set jsText to jsText & "else if(CAP.haveSig&&CAP.rssi<=-72)h+='<div class=\"advice bad\">Signal is weak at this position. Repeat the capture standing next to the access point. If the numbers improve sharply this is a coverage or placement problem, not a device problem.</div>';" & linefeed
set jsText to jsText & "else if(CAP.haveSig&&CAP.snr<25)h+='<div class=\"advice warn\">Signal strength is acceptable but the noise floor is high, which is the classic signature of a shared multi-tenant building.</div>';" & linefeed
set jsText to jsText & "else h+='<div class=\"advice good\">Signal and noise are within healthy limits at this position.</div>';" & linefeed
set jsText to jsText & "return h;}," & linefeed
set jsText to jsText & "rSignal_after:function(){var cv=document.getElementById('cvSig');if(cv)Chart.signal(cv);}," & linefeed
set jsText to jsText & "rChan:function(){var h='<canvas id=\"cvChan\"></canvas>';var same=0;CAP.nb.forEach(function(x){if(x.ch===CHAN)same=x.n;});" & linefeed
set jsText to jsText & "h+='<div class=\"result\"><div class=\"line\"><span>Neighbouring networks</span><b class=\"'+(CAP.nbTotal>12?'c-red':(CAP.nbTotal>6?'c-amber':'c-green'))+'\">'+CAP.nbTotal+'</b></div><div class=\"line\"><span>Sharing your channel ('+CHAN+')</span><b class=\"'+(same>2?'c-red':'')+'\">'+same+'</b></div><div class=\"line\"><span>Your band</span><b>'+BAND+'</b></div></div>';" & linefeed
set jsText to jsText & "if(CAP.nbTotal>12)h+='<div class=\"advice bad\">Dense RF environment. In a shared building this is a leading cause of intermittent Wi-Fi. Move the office SSID to a clear 5 GHz or 6 GHz channel and disable low legacy data rates.</div>';" & linefeed
set jsText to jsText & "else if(BAND==='2.4 GHz')h+='<div class=\"advice warn\">The client is associated on 2.4 GHz. Steer these Macs to 5 GHz. 2.4 GHz has only three non-overlapping channels and is shared with every neighbour, microwave and Bluetooth device nearby.</div>';" & linefeed
set jsText to jsText & "else h+='<div class=\"advice good\">Channel occupancy at this position looks reasonable.</div>';" & linefeed
set jsText to jsText & "return h;}," & linefeed
set jsText to jsText & "rChan_after:function(){var cv=document.getElementById('cvChan');if(cv)Chart.chan(cv);}," & linefeed
set jsText to jsText & "rCompare:function(){var h='<canvas id=\"cvCmp\"></canvas>';" & linefeed
set jsText to jsText & "if(CAPS.length<2)return h+'<div class=\"empty\">Only this capture is stored.<br>Run the tool again at another position - desk, next to the access point, then on wired Ethernet - and they will line up here automatically.</div>';" & linefeed
set jsText to jsText & "h+='<table><thead><tr><th>Capture point</th><th>Link</th><th style=\"text-align:right\">Loss</th><th style=\"text-align:right\">Avg</th><th style=\"text-align:right\">Jitter</th><th style=\"text-align:right\">RSSI</th><th style=\"text-align:right\">SNR</th><th style=\"text-align:right\">Drop</th></tr></thead><tbody>';" & linefeed
set jsText to jsText & "CAPS.slice().reverse().forEach(function(x){var cls=x.id===CAP.id?'best':(x.loss>=5?'bad':'');var d=new Date(x.ts);" & linefeed
set jsText to jsText & "h+='<tr class=\"'+cls+'\"><td><b>'+esc(x.loc)+'</b>'+(x.id===CAP.id?' <span class=\"badge new\">THIS</span>':'')+'<br><span class=\"mute\" style=\"font-size:9px\">'+d.toLocaleDateString()+' '+d.toLocaleTimeString([],{hour:'2-digit',minute:'2-digit'})+'</span></td><td class=\"dim\">'+esc(x.link)+'</td><td style=\"text-align:right\" class=\"'+(x.loss>0?'c-red':'c-green')+'\"><b>'+x.loss+'%</b></td><td style=\"text-align:right\">'+x.avg+' ms</td><td style=\"text-align:right\">'+x.jit+' ms</td><td style=\"text-align:right\">'+(x.haveSig?x.rssi:'--')+'</td><td style=\"text-align:right\">'+(x.haveSig?x.snr:'--')+'</td><td style=\"text-align:right\" class=\"'+(x.drop>=3?'c-red':'')+'\">'+x.drop+'</td></tr>';});" & linefeed
set jsText to jsText & "h+='</tbody></table>';" & linefeed
set jsText to jsText & "var wired=CAPS.filter(function(x){return x.link.indexOf('Wi-Fi')!==0;});" & linefeed
set jsText to jsText & "var wifi=CAPS.filter(function(x){return x.link.indexOf('Wi-Fi')===0;});" & linefeed
set jsText to jsText & "if(wired.length&&wifi.length){var ww=Math.max.apply(null,wired.map(function(x){return x.loss;}));var fw=Math.max.apply(null,wifi.map(function(x){return x.loss;}));" & linefeed
set jsText to jsText & "if(ww===0&&fw>0)h+='<div class=\"advice bad\"><b>Wired is clean, Wi-Fi is not.</b> The fault is wireless or environmental - access point coverage, RF interference or channel congestion. The circuit and the Macs are exonerated.</div>';" & linefeed
set jsText to jsText & "else if(ww>0)h+='<div class=\"advice bad\"><b>Loss is present on wired as well.</b> This is upstream of the wireless network. Escalate to the ISP and the office router with these captures attached.</div>';" & linefeed
set jsText to jsText & "else h+='<div class=\"advice good\">Both wired and wireless captures are clean. Repeat during the reported failure window with a Deep capture.</div>';}" & linefeed
set jsText to jsText & "else h+='<div class=\"advice warn\">You have not captured both a wireless and a wired sample yet. That single comparison is what separates a Wi-Fi problem from an ISP problem. Plug in an Ethernet adapter and run the Wired Ethernet Test.</div>';" & linefeed
set jsText to jsText & "return h+'<button class=\"btn ghost sm noprint\" style=\"margin-top:9px\" data-act=\"clearLedger\">CLEAR CAPTURE LEDGER</button>';}," & linefeed
set jsText to jsText & "rCompare_after:function(){var cv=document.getElementById('cvCmp');if(cv)Chart.compare(cv);}," & linefeed
set jsText to jsText & "rPath:function(){var rows=[['Default gateway',CAP.router,CAP.gw==='clean'?'ok':(CAP.gw==='lossy'?'warn':'bad'),CAP.gw==='clean'?'0% loss':(CAP.gw==='lossy'?'responding with loss':'no response')],['DNS resolution',CAP.dns,CAP.dnsOk===1?'ok':'bad',CAP.dnsOk===1?'answered':'no answer'],['HTTPS reachability','apple.com test endpoint',CAP.httpCode==='200'?'ok':'bad','HTTP '+CAP.httpCode],['First traceroute hop',CAP.router,CAP.hop1===0?'ok':'warn',CAP.hop1===0?'answered':'silent'],['Internet echo',CAP.target,CAP.loss===0?'ok':(CAP.loss<5?'warn':'bad'),CAP.loss+'% loss']];" & linefeed
set jsText to jsText & "var h='<table><thead><tr><th>Check</th><th>Target</th><th>Result</th></tr></thead><tbody>';" & linefeed
set jsText to jsText & "rows.forEach(function(r){var pill=r[2]==='ok'?'<span class=\"badge ok\">PASS</span>':(r[2]==='warn'?'<span class=\"badge\">WATCH</span>':'<span class=\"badge bad\">FAIL</span>');" & linefeed
set jsText to jsText & "h+='<tr class=\"'+(r[2]==='bad'?'bad':'')+'\"><td><b>'+esc(r[0])+'</b></td><td class=\"dim\">'+esc(r[1])+'</td><td>'+pill+' <span class=\"mute\">'+esc(r[3])+'</span></td></tr>';});" & linefeed
set jsText to jsText & "return h+'</tbody></table><div class=\"hint\" style=\"margin-top:9px\">Reading order: if the gateway passes but the internet echo fails, the break is beyond the office router. If the gateway itself fails, stay inside the office.</div>';}," & linefeed
set jsText to jsText & "rActions:function(){var a=[];" & linefeed
set jsText to jsText & "if(!ISWIFI&&CAP.loss>0)a.push('Loss on a wired link. Collect these results and open a circuit ticket with the ISP, including the traceroute output below.');" & linefeed
set jsText to jsText & "if(ISWIFI&&CAP.loss>0)a.push('Repeat this capture on wired Ethernet from the same desk. That single test decides whether this is a wireless problem or a circuit problem.');" & linefeed
set jsText to jsText & "if(ISWIFI&&CAP.haveSig&&CAP.rssi<=-72)a.push('Repeat the capture standing next to the access point and compare the two entries in the comparison panel.');" & linefeed
set jsText to jsText & "if(CAP.nbTotal>12)a.push('Document the neighbouring network count for the landlord or network vendor. It is the evidence for a channel plan change.');" & linefeed
set jsText to jsText & "if(BAND==='2.4 GHz')a.push('Move the affected Macs onto the 5 GHz SSID, or enable band steering if the access point supports it.');" & linefeed
set jsText to jsText & "if(CAP.drop>=3)a.push('Check access point roaming and coverage overlap. The client is losing association, not just slowing down.');" & linefeed
set jsText to jsText & "if(CAP.dnsOk===0)a.push('Validate the DNS servers on the active interface before assuming a connectivity fault.');" & linefeed
set jsText to jsText & "if(CAP.n<60)a.push('Run an Extended or Deep capture during the reported failure window before drawing a conclusion.');" & linefeed
set jsText to jsText & "a.push('Capture the same three points on every affected machine: desk, near access point, wired. Consistent results across users point at the environment, not the endpoints.');" & linefeed
set jsText to jsText & "a.push('Attach the exported JSON and this HTML report to the ServiceNow incident as evidence.');" & linefeed
set jsText to jsText & "return '<ul style=\"margin:0;padding-left:18px\">'+a.map(function(x){return '<li style=\"margin:7px 0;line-height:1.55\">'+esc(x)+'</li>';}).join('')+'</ul>';}," & linefeed
set jsText to jsText & "rDevice:function(){return '<div class=\"result\" style=\"margin-top:0\"><div class=\"line\"><span>Computer</span><b>'+esc(CAP.host)+'</b></div><div class=\"line\"><span>Console user</span><b>'+esc(CAP.user)+'</b></div><div class=\"line\"><span>Model</span><b>'+esc(CAP.model)+' ('+esc(CAP.modelId)+')</b></div><div class=\"line\"><span>Chip</span><b>'+esc(CAP.chip)+'</b></div><div class=\"line\"><span>Serial</span><b>'+esc(CAP.serial)+'</b></div><div class=\"line\"><span>Operating system</span><b>'+esc(CAP.osName)+' '+esc(CAP.osVer)+' ('+esc(CAP.osBuild)+')</b></div><div class=\"line\"><span>Active interface</span><b>'+esc(CAP.iface)+' &middot; '+esc(CAP.link)+'</b></div><div class=\"line\"><span>Wi-Fi interface</span><b>'+esc(CAP.wifiDev)+'</b></div><div class=\"line\"><span>IP address</span><b>'+esc(CAP.ip)+'</b></div><div class=\"line\"><span>Default gateway</span><b>'+esc(CAP.router)+'</b></div><div class=\"line\"><span>DNS servers</span><b>'+esc(CAP.dns)+'</b></div><div class=\"line\"><span>Capture point</span><b class=\"c-cyan\">'+esc(CAP.loc)+'</b></div><div class=\"line\"><span>Output folder</span><b style=\"font-size:10px\">'+esc(CAP.out)+'</b></div></div>';}," & linefeed
set jsText to jsText & "rSamples:function(){var d=CAP.samples.slice();if(WIN>0&&d.length>WIN)d=d.slice(d.length-WIN);d=d.reverse();" & linefeed
set jsText to jsText & "var h='<table><thead><tr><th>#</th><th>Time</th><th>Status</th><th style=\"text-align:right\">Latency</th></tr></thead><tbody>';" & linefeed
set jsText to jsText & "d.slice(0,120).forEach(function(s){h+='<tr class=\"'+(s.ok?'':'bad')+'\"><td class=\"mute\">'+s.i+'</td><td class=\"dim\">'+esc(s.t)+'</td><td>'+(s.ok?'<span class=\"c-green\">reply</span>':'<span class=\"c-red\">timeout</span>')+'</td><td style=\"text-align:right\"><b>'+(s.ok?s.ms+' ms':'--')+'</b></td></tr>';});" & linefeed
set jsText to jsText & "return h+'</tbody></table><div class=\"hint\" style=\"margin-top:8px\">Newest first. The complete sample set is in PingSamples.csv in the output folder.</div>';}," & linefeed
set jsText to jsText & "rCopilot:function(){" & linefeed
set jsText to jsText & "var p1='Act as a senior enterprise network engineer supporting executives. Analyse this macOS Wi-Fi capture and rank the root cause. Capture point: '+CAP.loc+'. Link: '+CAP.link+'. Device: '+CAP.model+' '+CAP.chip+', '+CAP.osName+' '+CAP.osVer+'. SSID '+CAP.ssid+', channel '+CAP.chan+' ('+BAND+'), PHY '+CAP.phy+', transmit rate '+CAP.tx+'. RSSI '+CAP.rssi+' dBm, noise '+CAP.noise+' dBm, SNR '+CAP.snr+' dB. Neighbouring networks '+CAP.nbTotal+'. Ping to '+CAP.target+': '+CAP.n+' samples, '+CAP.ok+' replies, '+CAP.bad+' timeouts, '+CAP.loss+' percent loss, latency min avg max '+CAP.min+' '+CAP.avg+' '+CAP.max+' ms, jitter '+CAP.jit+' ms, longest unbroken outage '+CAP.drop+' samples. Gateway '+CAP.router+' status '+CAP.gw+'. DNS answered: '+(CAP.dnsOk===1?'yes':'no')+'. HTTPS status '+CAP.httpCode+'. Give me: 1) a ranked root cause analysis with the evidence for each, covering RF interference, access point coverage, channel congestion, roaming, local router, DNS, ISP and endpoint; 2) the exact next tests to run onsite; 3) what to tell the landlord or network vendor if this turns out to be environmental.';" & linefeed
set jsText to jsText & "var p2='Using the capture data below, write a ServiceNow incident update for Executive Support. Include an Executive Impact and Business Value statement, a technical Work Notes section listing the investigation performed and the measurements taken, a Resolution or Action Plan with clear next steps, and a Status line. Then write a short, warm, non-technical message for the executive assistant explaining what was found and what happens next, with no IT jargon. Capture: point '+CAP.loc+', link '+CAP.link+', packet loss '+CAP.loss+' percent, average latency '+CAP.avg+' ms, jitter '+CAP.jit+' ms, longest outage '+CAP.drop+' samples, RSSI '+CAP.rssi+' dBm, SNR '+CAP.snr+' dB, '+CAP.nbTotal+' neighbouring networks, channel '+CAP.chan+', gateway status '+CAP.gw+', top ranked cause '+SCORES[0].k+'.';" & linefeed
set jsText to jsText & "window.__P1=p1;window.__P2=p2;" & linefeed
set jsText to jsText & "return '<div class=\"hint\" style=\"margin-bottom:9px\">Pre-filled with the measurements from this capture. Paste straight into Copilot.</div><div class=\"row noprint\" style=\"margin-bottom:9px\"><button class=\"btn\" data-act=\"copy1\">COPY ROOT CAUSE PROMPT</button><button class=\"btn\" data-act=\"copy2\">COPY SERVICENOW PROMPT</button></div><pre>'+esc(p1)+'</pre><div style=\"height:9px\"></div><pre>'+esc(p2)+'</pre>';}," & linefeed
set jsText to jsText & "rRaw:function(){var b=CAP.raw;var secs=[['Gateway test',b.gw],['DNS resolution',b.dns],['HTTPS timing',b.http],['Traceroute',b.trace],['Routing table',b.route],['Interface detail',b.ifc],['scutil DNS',b.scutil],['Preferred wireless networks',b.pref],['system_profiler SPAirPortDataType',b.air]];" & linefeed
set jsText to jsText & "return secs.map(function(s){return '<div class=\"hint\" style=\"margin:10px 0 4px\">'+esc(s[0]).toUpperCase()+'</div><pre>'+esc(s[1]||'(no output)')+'</pre>';}).join('');}};" & linefeed
set jsText to jsText & "var Palette={items:[]," & linefeed
set jsText to jsText & "open:function(){this.items=[];var self=this;" & linefeed
set jsText to jsText & "Panels.defs.forEach(function(d){self.items.push({t:'Go to '+d.title,s:'PANEL',f:function(){Panels.focus(d.id);}});});" & linefeed
set jsText to jsText & "this.items.push({t:'Toggle dark / light theme',s:'T',f:function(){App.theme();}});" & linefeed
set jsText to jsText & "this.items.push({t:'Export capture snapshot (JSON)',s:'E',f:function(){App.exportJson();}});" & linefeed
set jsText to jsText & "this.items.push({t:'Copy root cause prompt',s:'C',f:function(){App.copy(window.__P1,'Root cause prompt copied');}});" & linefeed
set jsText to jsText & "this.items.push({t:'Copy ServiceNow prompt',s:'',f:function(){App.copy(window.__P2,'ServiceNow prompt copied');}});" & linefeed
set jsText to jsText & "this.items.push({t:'Print / save as PDF',s:'P',f:function(){window.print();}});" & linefeed
set jsText to jsText & "this.items.push({t:'Clear capture ledger',s:'',f:function(){Ledger.clear();setTimeout(function(){location.reload();},400);}});" & linefeed
set jsText to jsText & "this.items.push({t:'Keyboard shortcuts',s:'?',f:function(){App.help();}});" & linefeed
set jsText to jsText & "q('#palette').classList.add('show');q('#cmdInput').value='';this.list('');setTimeout(function(){q('#cmdInput').focus();},40);}," & linefeed
set jsText to jsText & "close:function(){q('#palette').classList.remove('show');}," & linefeed
set jsText to jsText & "list:function(f){f=(f||'').toLowerCase();var self=this;var r=this.items.filter(function(x){return x.t.toLowerCase().indexOf(f)>=0;});" & linefeed
set jsText to jsText & "q('#cmdList').innerHTML=r.map(function(x,i){return '<div class=\"cmd-item'+(i===0?' sel':'')+'\" data-act=\"cmd\" data-i=\"'+self.items.indexOf(x)+'\">'+esc(x.t)+'<span class=\"cs\">'+esc(x.s)+'</span></div>';}).join('')||'<div class=\"empty\">No matches</div>';}," & linefeed
set jsText to jsText & "runSel:function(){var el=q('#cmdList .cmd-item.sel');if(!el)return;var i=parseInt(el.getAttribute('data-i'),10);this.close();this.items[i].f();}," & linefeed
set jsText to jsText & "move:function(dir){var list=Array.prototype.slice.call(document.querySelectorAll('#cmdList .cmd-item'));if(!list.length)return;var i=0;list.forEach(function(x,k){if(x.classList.contains('sel'))i=k;});list[i].classList.remove('sel');var n=i+dir;if(n<0)n=list.length-1;if(n>=list.length)n=0;list[n].classList.add('sel');list[n].scrollIntoView({block:'nearest'});}};" & linefeed
set jsText to jsText & "var App={" & linefeed
set jsText to jsText & "theme:function(){var h=document.documentElement;var cur=h.getAttribute('data-theme')==='light'?'dark':'light';h.setAttribute('data-theme',cur);LS.set('theme',cur);Panels.renderAll();}," & linefeed
set jsText to jsText & "help:function(){q('#helpModal').classList.add('show');}," & linefeed
set jsText to jsText & "closeHelp:function(){q('#helpModal').classList.remove('show');}," & linefeed
set jsText to jsText & "full:function(){if(document.fullscreenElement)document.exitFullscreen();else document.documentElement.requestFullscreen();}," & linefeed
set jsText to jsText & "copy:function(t,m){try{navigator.clipboard.writeText(t);toast(m||'Copied','ok');}catch(e){toast('Copy was blocked by the browser','err');}}," & linefeed
set jsText to jsText & "exportJson:function(){var blob=new Blob([JSON.stringify({capture:CAP,scores:SCORES,ledger:CAPS},null,2)],{type:'application/json'});var a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='WiFiCapture_'+CAP.id+'.json';a.click();toast('Snapshot exported','ok');}," & linefeed
set jsText to jsText & "setWin:function(v){WIN=v;LS.set('win',v);Array.prototype.slice.call(document.querySelectorAll('#rangePills .pill')).forEach(function(p){p.classList.toggle('on',parseInt(p.getAttribute('data-win'),10)===v);});Panels.render('latency');Panels.render('samples');}," & linefeed
set jsText to jsText & "hero:function(){" & linefeed
set jsText to jsText & "q('#lamp').className='statuslamp '+VERDICT.lamp;q('#lamp').textContent=VERDICT.emo;" & linefeed
set jsText to jsText & "q('#heroTitle').textContent=VERDICT.t;" & linefeed
set jsText to jsText & "q('#heroSub').textContent=CAP.ssid+'   '+CAP.chan+'   '+BAND+'   '+CAP.link+'   '+CAP.n+' samples';" & linefeed
set jsText to jsText & "q('#heroBig').textContent=CAP.loss+'%';" & linefeed
set jsText to jsText & "q('#heroBig').style.color=CAP.loss>=5?'var(--red)':(CAP.loss>0?'var(--amber)':'var(--green)');" & linefeed
set jsText to jsText & "var hp=q('#heroPill');hp.className='statuspill '+VERDICT.pill;hp.textContent=VERDICT.p;" & linefeed
set jsText to jsText & "var lp=q('#locPill');lp.textContent=CAP.loc.toUpperCase();" & linefeed
set jsText to jsText & "var kp=q('#linkPill');kp.textContent=CAP.link.toUpperCase();kp.className='statuspill '+(ISWIFI?'sp-warn':'sp-ok');" & linefeed
set jsText to jsText & "var d=new Date(CAP.ts);" & linefeed
set jsText to jsText & "q('#clockLocal').textContent=d.toLocaleTimeString([],{hour:'2-digit',minute:'2-digit',second:'2-digit'});" & linefeed
set jsText to jsText & "q('#clockMeta').textContent=d.toLocaleDateString([],{weekday:'short',month:'short',day:'numeric'})+'  CAPTURED';" & linefeed
set jsText to jsText & "var kpi=[['PACKET LOSS',CAP.loss+'%',CAP.bad+' of '+CAP.n+' failed',CAP.loss>0?'--red':'--green']," & linefeed
set jsText to jsText & "['AVG LATENCY',CAP.avg+' ms','min '+CAP.min+' / max '+CAP.max,CAP.avg>90?'--amber':'--cyan']," & linefeed
set jsText to jsText & "['JITTER',CAP.jit+' ms','mean deviation',CAP.jit>30?'--amber':'--cyan']," & linefeed
set jsText to jsText & "['LONGEST DROP',CAP.drop+'','consecutive timeouts',CAP.drop>=3?'--red':'--green']," & linefeed
set jsText to jsText & "['RSSI',CAP.haveSig?CAP.rssi+' dBm':'--',CAP.haveSig?(CAP.rssi>-67?'strong':(CAP.rssi>-72?'acceptable':'weak')):'not available',CAP.haveSig?(CAP.rssi>-72?'--green':'--red'):'--txt-mute']," & linefeed
set jsText to jsText & "['SNR',CAP.haveSig?CAP.snr+' dB':'--',CAP.haveSig?(CAP.snr>=25?'clean':'noisy'):'not available',CAP.haveSig?(CAP.snr>=25?'--green':'--amber'):'--txt-mute']," & linefeed
set jsText to jsText & "['NEARBY NETWORKS',CAP.nbTotal+'','RF density',CAP.nbTotal>12?'--red':(CAP.nbTotal>6?'--amber':'--green')]," & linefeed
set jsText to jsText & "['GATEWAY',CAP.gw.toUpperCase(),CAP.router,CAP.gw==='clean'?'--green':'--red']];" & linefeed
set jsText to jsText & "q('#kpis').innerHTML=kpi.map(function(k){return '<div class=\"kpi\"><div class=\"k\">'+k[0]+'</div><div class=\"v\" style=\"color:var('+k[3]+')\">'+esc(k[1])+'</div><div class=\"d\">'+esc(k[2])+'</div></div>';}).join('');" & linefeed
set jsText to jsText & "var ab=q('#alertBanner');" & linefeed
set jsText to jsText & "if(CAP.loss>=5||CAP.drop>=3){ab.classList.add('show');q('#alertText').textContent='LINK DEGRADED - '+CAP.loss+'% packet loss, longest outage '+CAP.drop+' samples. Top ranked cause: '+SCORES[0].k+'.';}" & linefeed
set jsText to jsText & "else if(CAP.loss>0||(CAP.haveSig&&CAP.rssi<=-72)){ab.classList.add('show');q('#alertText').textContent='MARGINAL LINK - review the root cause matrix and repeat this capture at a second position.';}" & linefeed
set jsText to jsText & "else ab.classList.remove('show');}," & linefeed
set jsText to jsText & "init:function(){var th=LS.get('theme','dark');if(th==='light')document.documentElement.setAttribute('data-theme','light');" & linefeed
set jsText to jsText & "this.hero();Panels.chips();Panels.build();this.setWin(WIN);" & linefeed
set jsText to jsText & "toast('Capture stored. '+CAPS.length+' capture'+(CAPS.length===1?'':'s')+' in the comparison ledger.','ok',5200);}};" & linefeed
set jsText to jsText & "document.addEventListener('click',function(e){" & linefeed
set jsText to jsText & "if(e.target.id==='palette'){Palette.close();return;}" & linefeed
set jsText to jsText & "if(e.target.id==='helpModal'){App.closeHelp();return;}" & linefeed
set jsText to jsText & "var t=e.target.closest('[data-act]');if(!t)return;" & linefeed
set jsText to jsText & "var a=t.getAttribute('data-act'),id=t.getAttribute('data-id');" & linefeed
set jsText to jsText & "if(a==='palette')Palette.open();" & linefeed
set jsText to jsText & "else if(a==='theme')App.theme();" & linefeed
set jsText to jsText & "else if(a==='rerender'){Panels.renderAll();toast('Panels re-rendered','ok',2200);}" & linefeed
set jsText to jsText & "else if(a==='export')App.exportJson();" & linefeed
set jsText to jsText & "else if(a==='print')window.print();" & linefeed
set jsText to jsText & "else if(a==='full')App.full();" & linefeed
set jsText to jsText & "else if(a==='dismiss')q('#alertBanner').classList.remove('show');" & linefeed
set jsText to jsText & "else if(a==='chip')Panels.toggle(id);" & linefeed
set jsText to jsText & "else if(a==='collapse')Panels.collapse(id);" & linefeed
set jsText to jsText & "else if(a==='hide'){e.stopPropagation();Panels.toggle(id);}" & linefeed
set jsText to jsText & "else if(a==='closeHelp')App.closeHelp();" & linefeed
set jsText to jsText & "else if(a==='cmd')Palette.runSel();" & linefeed
set jsText to jsText & "else if(a==='copy1')App.copy(window.__P1,'Root cause prompt copied');" & linefeed
set jsText to jsText & "else if(a==='copy2')App.copy(window.__P2,'ServiceNow prompt copied');" & linefeed
set jsText to jsText & "else if(a==='clearLedger'){Ledger.clear();toast('Capture ledger cleared','warn');setTimeout(function(){location.reload();},600);}});" & linefeed
set jsText to jsText & "document.addEventListener('mouseover',function(e){var t=e.target.closest('.cmd-item');if(!t)return;Array.prototype.slice.call(document.querySelectorAll('#cmdList .cmd-item')).forEach(function(x){x.classList.remove('sel');});t.classList.add('sel');});" & linefeed
set jsText to jsText & "q('#cmdInput').addEventListener('input',function(){Palette.list(this.value);});" & linefeed
set jsText to jsText & "Array.prototype.slice.call(document.querySelectorAll('#rangePills .pill')).forEach(function(p){p.addEventListener('click',function(){App.setWin(parseInt(p.getAttribute('data-win'),10));});});" & linefeed
set jsText to jsText & "document.addEventListener('keydown',function(e){" & linefeed
set jsText to jsText & "var pal=q('#palette').classList.contains('show');" & linefeed
set jsText to jsText & "if((e.metaKey||e.ctrlKey)&&e.key.toLowerCase()==='k'){e.preventDefault();if(pal){Palette.close();}else{Palette.open();}return;}" & linefeed
set jsText to jsText & "if(pal){if(e.key==='Escape')Palette.close();if(e.key==='ArrowDown'){e.preventDefault();Palette.move(1);}if(e.key==='ArrowUp'){e.preventDefault();Palette.move(-1);}if(e.key==='Enter'){e.preventDefault();Palette.runSel();}return;}" & linefeed
set jsText to jsText & "if(e.target.tagName==='INPUT')return;" & linefeed
set jsText to jsText & "var k=e.key.toLowerCase();" & linefeed
set jsText to jsText & "if(e.key==='Escape')App.closeHelp();" & linefeed
set jsText to jsText & "else if(k==='t')App.theme();" & linefeed
set jsText to jsText & "else if(k==='f')App.full();" & linefeed
set jsText to jsText & "else if(k==='r')Panels.renderAll();" & linefeed
set jsText to jsText & "else if(k==='e')App.exportJson();" & linefeed
set jsText to jsText & "else if(k==='p')window.print();" & linefeed
set jsText to jsText & "else if(k==='c')App.copy(window.__P1,'Root cause prompt copied');" & linefeed
set jsText to jsText & "else if(k==='?')App.help();" & linefeed
set jsText to jsText & "else if(k==='1')App.setWin(25);" & linefeed
set jsText to jsText & "else if(k==='2')App.setWin(50);" & linefeed
set jsText to jsText & "else if(k==='3')App.setWin(100);" & linefeed
set jsText to jsText & "else if(k==='4')App.setWin(0);});" & linefeed
set jsText to jsText & "window.addEventListener('resize',function(){Panels.renderAll();});" & linefeed
set jsText to jsText & "App.init();"

-- ------------------------------------------------------------ write report ---

my writeNew(htmlPath, "<!DOCTYPE html><html lang='en'><head><meta charset='UTF-8'>")
my appendText(htmlPath, "<meta name='viewport' content='width=device-width, initial-scale=1.0, viewport-fit=cover'>")
my appendText(htmlPath, "<meta name='theme-color' content='#05080d'>")
my appendText(htmlPath, "<title>WiFi Connectivity Monitor - " & hostName & "</title><style>")
my appendText(htmlPath, cssText)
my appendText(htmlPath, shellText)
my appendText(htmlPath, "var CAP=" & capJs & ";" & linefeed)
my appendText(htmlPath, jsText)
my appendText(htmlPath, linefeed & "</scr" & "ipt></body></html>")

my appendText(logPath, "[" & (do shell script "date '+%Y-%m-%d %H:%M:%S'") & "] Report written: " & htmlPath & linefeed)

do shell script "open " & quoted form of htmlPath

set sigLine to "not available"
if haveSignal then set sigLine to (rssiVal as text) & " dBm / SNR " & (snrVal as text) & " dB"

set summaryText to "WiFi Connectivity Monitor v3.0 complete."
set summaryText to summaryText & return & return & "Capture point:  " & captureLocation
set summaryText to summaryText & return & "Link:  " & linkType
set summaryText to summaryText & return & "Packet loss:  " & (lossPct as text) & " percent"
set summaryText to summaryText & return & "Avg latency:  " & (avgLat as text) & " ms"
set summaryText to summaryText & return & "Jitter:  " & (jitterVal as text) & " ms"
set summaryText to summaryText & return & "Longest drop:  " & (maxConsecFail as text) & " samples"
set summaryText to summaryText & return & "Signal:  " & sigLine
set summaryText to summaryText & return & return & "Report folder:" & return & outDir

display dialog summaryText buttons {"OK"} default button "OK" with title "WiFi Connectivity Monitor v3.0"
