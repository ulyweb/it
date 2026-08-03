-- WiFi Connectivity Diagnostic Capture - macOS Script Editor App
-- Purpose: Read-only Executive Support network telemetry capture for Wi-Fi drops, RF/signal issues,
-- gateway/ISP validation, DNS testing, ping stability, and HTML dashboard reporting.
-- Save as Application from Script Editor for desk-side use.

set scriptBody to "#!/bin/bash
set -u
TS=$(date +%Y%m%d_%H%M%S)
BASE=\"/Users/Shared/WiFiConnectivityDiag\"
OUT=\"$BASE/reports/${TS}_WiFiConnectivityDiag\"
mkdir -p \"$OUT\"
LOG=\"$OUT/WiFiConnectivityDiag.log\"
HTML=\"$OUT/WiFiConnectivityDiag_Report.html\"
PINGCSV=\"$OUT/PingSamples.csv\"
SYSTEMTXT=\"$OUT/SystemSummary.txt\"
WIFITXT=\"$OUT/WiFiDetails.txt\"
TARGET_IP=\"8.8.8.8\"
TARGET_DNS=\"google.com\"
PING_COUNT=60

log(){ echo \"[$(date '+%Y-%m-%d %H:%M:%S')] $*\" >> \"$LOG\"; }
esc(){ /usr/bin/python3 -c 'import html,sys; print(html.escape(sys.stdin.read()))'; }
run(){ log \"RUN: $*\"; bash -lc \"$*\" 2>&1; }

log \"Starting macOS WiFi Connectivity Diagnostic Capture\"

HOST=$(scutil --get ComputerName 2>/dev/null || hostname)
USER_NAME=$(stat -f %Su /dev/console 2>/dev/null || whoami)
SERIAL=$(system_profiler SPHardwareDataType 2>/dev/null | awk -F': ' '/Serial Number/{print $2; exit}')
MODEL=$(system_profiler SPHardwareDataType 2>/dev/null | awk -F': ' '/Model Name/{print $2; exit}')
CHIP=$(system_profiler SPHardwareDataType 2>/dev/null | awk -F': ' '/Chip|Processor Name/{print $2; exit}')
MACOS=$(sw_vers -productVersion 2>/dev/null)
BUILD=$(sw_vers -buildVersion 2>/dev/null)
WIFI_DEVICE=$(networksetup -listallhardwareports 2>/dev/null | awk '/Wi-Fi|AirPort/{getline; print $2; exit}')
[ -z \"${WIFI_DEVICE:-}\" ] && WIFI_DEVICE=\"en0\"

AIRPORT=\"/System/Library/PrivateFrameworks/Apple80211.framework/Versions/Current/Resources/airport\"
WIFI_INFO=\"airport utility not available\"
if [ -x \"$AIRPORT\" ]; then
  WIFI_INFO=$(\"$AIRPORT\" -I 2>&1)
fi

SSID=$(echo \"$WIFI_INFO\" | awk -F': ' '/ SSID/{print $2; exit}')
BSSID=$(echo \"$WIFI_INFO\" | awk -F': ' '/ BSSID/{print $2; exit}')
RSSI=$(echo \"$WIFI_INFO\" | awk -F': ' '/agrCtlRSSI/{print $2; exit}')
NOISE=$(echo \"$WIFI_INFO\" | awk -F': ' '/agrCtlNoise/{print $2; exit}')
CHANNEL=$(echo \"$WIFI_INFO\" | awk -F': ' '/ channel/{print $2; exit}')
TXRATE=$(echo \"$WIFI_INFO\" | awk -F': ' '/lastTxRate/{print $2; exit}')
PHY=$(echo \"$WIFI_INFO\" | awk -F': ' '/ PHY Mode/{print $2; exit}')
SNR=\"N/A\"
if [[ \"$RSSI\" =~ ^-?[0-9]+$ && \"$NOISE\" =~ ^-?[0-9]+$ ]]; then SNR=$((RSSI-NOISE)); fi

ROUTER=$(route -n get default 2>/dev/null | awk '/gateway:/{print $2; exit}')
IPADDR=$(ipconfig getifaddr \"$WIFI_DEVICE\" 2>/dev/null || true)
DNS_SERVERS=$(scutil --dns 2>/dev/null | awk '/nameserver\[[0-9]+\]/{print $3}' | sort -u | paste -sd ', ' -)

{
 echo \"Computer: $HOST\"
 echo \"Console User: $USER_NAME\"
 echo \"Model: $MODEL\"
 echo \"Chip: $CHIP\"
 echo \"Serial: $SERIAL\"
 echo \"macOS: $MACOS ($BUILD)\"
 echo \"Wi-Fi Device: $WIFI_DEVICE\"
 echo \"IP Address: $IPADDR\"
 echo \"Router: $ROUTER\"
 echo \"DNS Servers: $DNS_SERVERS\"
} > \"$SYSTEMTXT\"

{
 echo \"$WIFI_INFO\"
 echo
 echo \"networksetup -getairportnetwork $WIFI_DEVICE\"
 networksetup -getairportnetwork \"$WIFI_DEVICE\" 2>&1
 echo
 echo \"ifconfig $WIFI_DEVICE\"
 ifconfig \"$WIFI_DEVICE\" 2>&1
} > \"$WIFITXT\"

echo \"Sample,Timestamp,Target,Status,LatencyMs\" > \"$PINGCSV\"
SUCCESS=0
FAIL=0
SUM=0
MAX=0
for i in $(seq 1 $PING_COUNT); do
  TSNOW=$(date '+%Y-%m-%d %H:%M:%S')
  OUTPING=$(ping -c 1 -W 1000 \"$TARGET_IP\" 2>&1)
  LAT=$(echo \"$OUTPING\" | awk -F'time=' '/time=/{split($2,a,\" \" ); print int(a[1]); exit}')
  if [ -n \"$LAT\" ]; then
    echo \"$i,$TSNOW,$TARGET_IP,Success,$LAT\" >> \"$PINGCSV\"
    SUCCESS=$((SUCCESS+1)); SUM=$((SUM+LAT)); [ $LAT -gt $MAX ] && MAX=$LAT
  else
    echo \"$i,$TSNOW,$TARGET_IP,Failed,\" >> \"$PINGCSV\"
    FAIL=$((FAIL+1))
  fi
  sleep 1
done
LOSS=$(awk -v f=$FAIL -v c=$PING_COUNT 'BEGIN{printf \"%.2f\", (f/c)*100}')
AVG=\"N/A\"; [ $SUCCESS -gt 0 ] && AVG=$(awk -v s=$SUM -v c=$SUCCESS 'BEGIN{printf \"%.2f\", s/c}')
GATEWAY_STATUS=\"Not reachable / not detected\"
if [ -n \"${ROUTER:-}\" ] && ping -c 3 -W 1000 \"$ROUTER\" >/dev/null 2>&1; then GATEWAY_STATUS=\"Reachable\"; fi

DNS_TEST=$(run \"dig +time=2 +tries=1 $TARGET_DNS\")
TRACE_TEST=$(run \"traceroute -n -m 12 $TARGET_IP\")
NETSTAT_TEST=$(run \"netstat -rn\")
SCUTIL_TEST=$(run \"scutil --dns\")

HEALTH=\"GREEN\"; HEALTH_TEXT=\"No packet loss detected during sampling\"
LOSS_INT=${LOSS%.*}
if [ $LOSS_INT -gt 0 ] && [ $LOSS_INT -lt 10 ]; then HEALTH=\"AMBER\"; HEALTH_TEXT=\"Intermittent packet loss detected\"; fi
if [ $LOSS_INT -ge 10 ]; then HEALTH=\"RED\"; HEALTH_TEXT=\"Significant packet loss detected\"; fi

REC1=\"Run this same capture near the router/access point and again at the affected desk location to compare RSSI, SNR, latency, and packet loss.\"
REC2=\"If wired Ethernet is stable while Wi-Fi drops, focus on RF interference, access point placement, band/channel congestion, and office layout.\"
REC3=\"If wired and wireless both fail, focus on Comcast Business service, modem/router health, DHCP, DNS, and upstream packet loss.\"
REC4=\"For multi-tenant offices, review neighboring SSID congestion and confirm whether the affected area has weak signal or high noise.\"

STYLE='body{margin:0;background:#070707;color:#f5f2e9;font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Arial,sans-serif}.bg{position:fixed;inset:0;background:radial-gradient(circle at 20% 10%,rgba(212,175,55,.18),transparent 30%),radial-gradient(circle at 80% 30%,rgba(212,175,55,.10),transparent 25%),linear-gradient(135deg,#080808,#131313 45%,#050505);z-index:-1}.wrap{max-width:1200px;margin:0 auto;padding:34px}.hero{border:1px solid rgba(212,175,55,.35);background:rgba(18,18,18,.72);backdrop-filter:blur(18px);border-radius:24px;padding:28px;box-shadow:0 0 40px rgba(212,175,55,.12)}h1{margin:0;color:#ffd76a;font-size:32px}.sub{color:#cfc6aa;margin-top:8px}.grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:16px;margin:22px 0}.card{border:1px solid rgba(212,175,55,.25);background:rgba(255,255,255,.05);border-radius:20px;padding:18px;box-shadow:0 14px 40px rgba(0,0,0,.25)}.k{color:#bca75a;font-size:12px;text-transform:uppercase;letter-spacing:1px}.v{font-size:24px;font-weight:700;margin-top:8px}.pill{display:inline-block;padding:6px 12px;border-radius:999px;font-weight:700}.GREEN{background:#073b24;color:#66f0a5}.AMBER{background:#4a3500;color:#ffd36b}.RED{background:#4b0b0b;color:#ff7b7b}pre{white-space:pre-wrap;background:#0d0d0d;border:1px solid rgba(212,175,55,.2);border-radius:16px;padding:16px;overflow:auto;color:#eee}.section{margin-top:24px}.section h2{color:#ffd76a;border-bottom:1px solid rgba(212,175,55,.25);padding-bottom:8px}.small{color:#bfb7a1;font-size:12px}li{margin:8px 0}.two{display:grid;grid-template-columns:1fr 1fr;gap:18px}@media(max-width:900px){.grid,.two{grid-template-columns:1fr}}'

{
cat <<EOF
<!doctype html><html><head><meta charset=\"utf-8\"><title>WiFi Connectivity Diagnostic Report - macOS</title><style>$STYLE</style></head><body><div class=\"bg\"></div><div class=\"wrap\">
<div class=\"hero\"><h1>macOS WiFi Connectivity Diagnostic Report</h1><div class=\"sub\">Executive Support telemetry capture for intermittent Wi-Fi, RF/signal, Comcast Business, gateway, DNS, and endpoint validation.</div><p><span class=\"pill $HEALTH\">$HEALTH</span> $HEALTH_TEXT</p></div>
<div class=\"grid\">
<div class=\"card\"><div class=\"k\">Computer</div><div class=\"v\">$(echo \"$HOST\" | esc)</div></div>
<div class=\"card\"><div class=\"k\">SSID</div><div class=\"v\">$(echo \"${SSID:-Not detected}\" | esc)</div></div>
<div class=\"card\"><div class=\"k\">Packet Loss</div><div class=\"v\">$LOSS%</div></div>
<div class=\"card\"><div class=\"k\">Avg / Max Latency</div><div class=\"v\">$AVG / $MAX ms</div></div>
</div>
<div class=\"two\">
<div class=\"card\"><div class=\"k\">Device</div><div>$(echo \"$MODEL $CHIP\" | esc)<br>Serial: $(echo \"$SERIAL\" | esc)<br>macOS: $(echo \"$MACOS ($BUILD)\" | esc)</div></div>
<div class=\"card\"><div class=\"k\">Wireless Signal</div><div>RSSI: $(echo \"${RSSI:-N/A}\" | esc) dBm<br>Noise: $(echo \"${NOISE:-N/A}\" | esc) dBm<br>SNR: $(echo \"$SNR\" | esc)<br>Channel: $(echo \"${CHANNEL:-N/A}\" | esc)<br>Tx Rate: $(echo \"${TXRATE:-N/A}\" | esc)<br>PHY: $(echo \"${PHY:-N/A}\" | esc)</div></div>
</div>
<div class=\"section card\"><h2>Recommended Next Actions</h2><ul><li>$REC1</li><li>$REC2</li><li>$REC3</li><li>$REC4</li></ul></div>
<div class=\"section card\"><h2>Network Summary</h2><pre>$(cat \"$SYSTEMTXT\" | esc)</pre></div>
<div class=\"section card\"><h2>Ping Samples</h2><p>Target: $TARGET_IP | Count: $PING_COUNT | Success: $SUCCESS | Failed: $FAIL | Gateway: $(echo \"${ROUTER:-Not detected}\" | esc) - $(echo \"$GATEWAY_STATUS\" | esc)</p><pre>$(tail -25 \"$PINGCSV\" | esc)</pre></div>
<div class=\"section card\"><h2>Wi-Fi Details</h2><pre>$(cat \"$WIFITXT\" | esc)</pre></div>
<div class=\"section card\"><h2>DNS Test</h2><pre>$(echo \"$DNS_TEST\" | esc)</pre></div>
<div class=\"section card\"><h2>Traceroute</h2><pre>$(echo \"$TRACE_TEST\" | esc)</pre></div>
<div class=\"section card\"><h2>Routing Table</h2><pre>$(echo \"$NETSTAT_TEST\" | esc)</pre></div>
<div class=\"section card\"><h2>scutil DNS</h2><pre>$(echo \"$SCUTIL_TEST\" | esc)</pre></div>
<p class=\"small\">Generated: $(date) | Output folder: $OUT</p>
</div></body></html>
EOF
} > \"$HTML\"

log \"Report generated: $HTML\"
open \"$HTML\"
echo \"$HTML\"
"

set tempScript to "/tmp/wifi_connectivity_diag_mac.sh"
do shell script "cat > " & quoted form of tempScript & " <<'EOS'" & linefeed & scriptBody & linefeed & "EOS" & linefeed & "chmod +x " & quoted form of tempScript
set reportPath to do shell script quoted form of tempScript

display dialog "WiFi Connectivity Diagnostic report has been generated." & return & return & reportPath buttons {"OK"} default button "OK" with title "WiFi Connectivity Diagnostic"
