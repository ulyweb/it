-- ============================================================================
--  TAHOE-DG Update Assistant  v1.4
--  macOS Tahoe & Digital Guardian 9.2 Compliance Assistant
--
--  Guided GUI tool for Mac updates: macOS Software Update +
--  Digital Guardian (via Jamf Self Service policy), with PROACTIVE compliance
--  pre-check, an OPEN-APPS safety gate, full audit logging, and a dark/gold
--  telemetry HTML receipt.
--
--  v1.1: Pre-flight compliance gate (macOS >= 26 Tahoe, DG >= 9.2).
--  v1.2: Open-apps safety gate (must close apps before update+restart).
--  v1.3: Multi-source DG detection.
--  v1.4: FIXED DG detection. Prior probe used fragile multi-line shell with
--        line-continuations + escaped regex that could throw at runtime and
--        silently fall back to "Unknown", so a COMPLIANT Mac (DGCIApp/dgdaemon
--        at 9.2.0.0056, visible in System Information) was wrongly sent to
--        Self Service. v1.4 parses `system_profiler SPApplicationsDataType`
--        directly for the DG components (DGCIApp/DGCipher/dgdaemon/dgesc),
--        validates the version with a simple, escaping-free pattern, LOGS the
--        detected value for audit, and only then decides. Compliance is always
--        evaluated BEFORE the open-apps scan and BEFORE any credential prompt.
--
--  Build:  Script Editor > paste > File > Export...
--          File Format: Application | Code Sign: Sign to Run Locally
--          Save to: /Applications  (or /Users/Shared)
--
--  Config: set DG_POLICY_EVENT to your Jamf custom trigger for Digital Guardian.
-- ============================================================================

property appTitle : "TAHOE-DG Update Assistant"
property appSubtitle : "macOS Tahoe & Digital Guardian 9.2 Compliance Assistant"
property appVersion : "1.4"
property MACOS_MIN_MAJOR : 26           -- Tahoe baseline
property DG_APPROVED_VERSION : "9.2"    -- corporate-approved DG baseline
property DG_POLICY_EVENT : "install_digitalguardian"   -- <-- set to your Jamf event trigger
property baseDir : "/Users/Shared/TAHOE-DG"

-- runtime globals
global adminUser, adminPass, outDir, logPath, htmlPath
global serialNum, hostName, osBefore, osBuild, dgBefore, consoleUser
global macCompliant, dgCompliant

on run
    set adminUser to ""
    set adminPass to ""
    
    -- ---- Environment detection ---------------------------------------------
    set consoleUser to do shell script "/usr/bin/stat -f%Su /dev/console"
    set serialNum to do shell script "system_profiler SPHardwareDataType 2>/dev/null | awk '/Serial Number/{print $NF; exit}'"
    if serialNum is "" then set serialNum to "UNKNOWN"
    set hostName to do shell script "scutil --get ComputerName 2>/dev/null || hostname"
    set osBefore to do shell script "sw_vers -productVersion"
    set osBuild to do shell script "sw_vers -buildVersion"
    
    -- ---- Output dir (self-healing) FIRST so we can log detection -----------
    set ts to do shell script "date '+%Y%m%d_%H%M%S'"
    set outDir to baseDir & "/" & ts & "_" & serialNum
    try
        do shell script "mkdir -p " & quoted form of outDir
    end try
    try
        do shell script "test -w " & quoted form of outDir
    on error
        set outDir to (system attribute "HOME") & "/TAHOE-DG/" & ts & "_" & serialNum
        do shell script "mkdir -p " & quoted form of outDir
    end try
    set logPath to outDir & "/update.log"
    set htmlPath to outDir & "/TAHOE-DG_Receipt_" & ts & ".html"
    do shell script "echo '' > " & quoted form of logPath
    my logLine("=== " & appTitle & " v" & appVersion & " started ===")
    my logLine(appSubtitle)
    
    -- ---- Version detection (with raw evidence logged) ----------------------
    set dgBefore to my detectDG()
    set macCompliant to my isMacCompliant(osBefore)
    set dgCompliant to my isDGCompliant(dgBefore)
    my logLine("Host " & hostName & " | Serial " & serialNum & " | macOS " & osBefore & " (" & osBuild & ")")
    my logLine("DETECTED Digital Guardian version: " & dgBefore)
    my logLine("Compliance: macOS " & osBefore & " >= " & MACOS_MIN_MAJOR & "? " & macCompliant & " | DG " & dgBefore & " >= " & DG_APPROVED_VERSION & "? " & dgCompliant)
    
    -- ---- Welcome / save-your-work warning ----------------------------------
    display dialog "Welcome to " & appTitle & "." & return & appSubtitle & return & return & "This will check your Mac and, if needed, update:" & return & "  • macOS Software Update  (~25–35 min, restart required)" & return & "  • Digital Guardian security agent  (~15 min, restart required)" & return & return & "⚠️ Please SAVE AND CLOSE all your work before continuing." buttons {"Cancel", "Continue"} default button "Continue" cancel button "Cancel" with title appTitle with icon note
    
    -- ---- Choose which updates ----------------------------------------------
    set taskList to {"Both  (macOS + Digital Guardian) — recommended", "macOS Software Update only", "Digital Guardian only"}
    set taskChoice to (choose from list taskList with prompt "Which updates would you like to check / run?" default items {"Both  (macOS + Digital Guardian) — recommended"} with title appTitle without empty selection allowed)
    if taskChoice is false then my bailOut("Cancelled at task selection.")
    set chosen to item 1 of taskChoice
    set doMac to (chosen does not contain "Digital Guardian only")
    set doDG to (chosen does not contain "macOS Software Update only")
    my logLine("Selection: macOS=" & doMac & " DigitalGuardian=" & doDG)
    
    -- ---- PROACTIVE compliance gate (ALWAYS before apps/credentials) --------
    set needMac to (doMac and not macCompliant)
    set needDG to (doDG and not dgCompliant)
    
    -- Safeguard: DG selected but version genuinely could not be determined
    if doDG and (dgBefore is "Unknown") then
        try
            set r to display dialog "⚠️  Digital Guardian version could not be determined" & return & return & "The tool checked this Mac but could not read an installed Digital Guardian version (baseline is " & DG_APPROVED_VERSION & "+)." & return & return & "If you know it is already current, choose Cancel. Otherwise you can install/repair via Self Service." buttons {"Cancel", "Proceed to Install"} default button "Cancel" cancel button "Cancel" with title (appTitle & " — DG Version Unknown") with icon caution
        on error number -128
            my bailOut("Cancelled — Digital Guardian version unknown, user chose not to proceed.")
        end try
    end if
    
    -- Case 1: nothing selected actually needs updating -> celebrate + exit
    if (doMac or doDG) and (not needMac) and (not needDG) then
        my logLine("All selected items already compliant. No action needed. (No apps scan, no credentials.)")
        my writeReceipt(doMac, doDG, "None needed", "Already up to date", "Already up to date", dgBefore)
        try
            do shell script "open " & quoted form of htmlPath
        end try
        set okMsg to "✅  You're all set — no updates needed!" & return & return & "This Mac already meets the current security baseline:" & return
        if doMac then set okMsg to okMsg & "  • macOS: " & osBefore & "  (Tahoe " & MACOS_MIN_MAJOR & "+ ✔)" & return
        if doDG then set okMsg to okMsg & "  • Digital Guardian: " & dgBefore & "  (" & DG_APPROVED_VERSION & "+ ✔)" & return
        set okMsg to okMsg & return & "No restart or further action is required. A compliance report was saved for your records."
        display dialog okMsg buttons {"Great, thanks!"} default button "Great, thanks!" with title (appTitle & " — Already Up To Date") with icon note giving up after 60
        return
    end if
    
    -- Case 2: partial compliance -> tell the user what will be skipped
    set skipMsg to ""
    if doMac and macCompliant then set skipMsg to skipMsg & "  • macOS is already current (" & osBefore & ") — will skip." & return
    if doDG and dgCompliant then set skipMsg to skipMsg & "  • Digital Guardian is already current (" & dgBefore & ") — will skip." & return
    if skipMsg is not "" then
        display dialog "Good news — some items are already up to date:" & return & return & skipMsg & return & "Only the remaining item(s) will be updated." buttons {"Continue"} default button "Continue" with title appTitle with icon note giving up after 45
    end if
    
    -- ---- OPEN-APPS SAFETY GATE (only because a restart is coming) ----------
    if not my ensureAppsClosed() then my bailOut("Update cancelled — open apps were not closed.")
    
    -- ---- Credential prompt (only because real work remains) ----------------
    if not my collectCredential() then my bailOut("Authentication failed or cancelled. No changes made.")
    my logLine("Admin credential validated for: " & adminUser)
    
    -- ---- Run only what is needed -------------------------------------------
    set macResult to "Skipped (already current)"
    set dgResult to "Skipped (already current)"
    set updatesFound to "n/a"
    
    if needDG then
        my logLine("---- Digital Guardian update ----")
        set dgResult to my runDigitalGuardian()
    else if doDG then
        my logLine("DG already compliant (" & dgBefore & ") — skipped.")
    end if
    
    if needMac then
        my logLine("---- macOS Software Update ----")
        set updatesFound to my listMacUpdates()
        set macResult to my installMacUpdates()
    else if doMac then
        my logLine("macOS already compliant (" & osBefore & ") — skipped.")
    end if
    
    set dgAfter to my detectDG()
    
    -- ---- Write telemetry receipt -------------------------------------------
    my writeReceipt(doMac, doDG, updatesFound, macResult, dgResult, dgAfter)
    my logLine("=== Completed. macOS=" & macResult & " | DG=" & dgResult & " ===")
    set adminPass to ""
    
    -- ---- Show summary + restart --------------------------------------------
    try
        do shell script "open " & quoted form of htmlPath
    end try
    
    set needsRestart to (needMac and macResult contains "installed") or (needDG and dgResult contains "policy")
    if needsRestart then
        try
            display dialog "Updates finished. A restart is required to complete installation." & return & return & "A detailed report was saved to:" & return & outDir & return & return & "Restart now?" buttons {"Restart Later", "Restart Now"} default button "Restart Now" with title appTitle with icon note giving up after 120
            if button returned of result is "Restart Now" then
                my logLine("User chose Restart Now.")
                do shell script "shutdown -r +1 'TAHOE-DG: restarting in 1 minute to finish updates.'" user name adminUser password adminPass with administrator privileges
                display dialog "Your Mac will restart in about 1 minute. Please save anything still open." buttons {"OK"} default button "OK" with title appTitle giving up after 20
            else
                my logLine("User deferred restart.")
            end if
        end try
    else
        display dialog "Finished. No restart was required." & return & return & "Report saved to:" & return & outDir buttons {"OK"} default button "OK" with title appTitle
    end if
end run

-- ============================ HANDLERS ======================================

-- RELIABLE Digital Guardian detection.
-- Parses system_profiler for DG components (DGCIApp/DGCipher/dgdaemon/dgesc),
-- validates the result with an escaping-free pattern, falls back to pkgutil.
-- Returns a real version string (e.g., 9.2.0.0056) or "Unknown".
on detectDG()
    set v to "Unknown"
    set raw to ""
    try
        set raw to do shell script "V=$(system_profiler SPApplicationsDataType 2>/dev/null | awk '/DGCIApp:|DGCipher:|dgdaemon:|dgesc:|DGNetworkExtensionManager:|Digital Guardian/{f=1} f&&/Version:/{print $2; exit}'); " & "if [ -z \"$V\" ]; then P=$(pkgutil --pkgs 2>/dev/null | grep -i guardian | head -1); if [ -n \"$P\" ]; then V=$(pkgutil --pkg-info \"$P\" 2>/dev/null | awk '/version:/{print $2; exit}'); fi; fi; " & "printf '%s' \"$V\""
    end try
    if raw is not "" then set v to raw
    -- Validate it looks like a version number (escaping-free pattern: [.] not \.)
    if v is not "Unknown" then
        set looksValid to true
        try
            do shell script "printf '%s' " & quoted form of v & " | grep -Eq '^[0-9]+[.][0-9]'"
        on error
            set looksValid to false
        end try
        if not looksValid then set v to "Unknown"
    end if
    try
        my logLine("detectDG() raw='" & raw & "' -> resolved='" & v & "'")
    end try
    return v
end detectDG

-- macOS compliant if major version >= MACOS_MIN_MAJOR (e.g., 26 = Tahoe)
on isMacCompliant(verStr)
    try
        set maj to (do shell script "echo " & quoted form of verStr & " | cut -d. -f1")
        if (maj as integer) ≥ MACOS_MIN_MAJOR then return true
    end try
    return false
end isMacCompliant

-- DG compliant if detected version >= approved baseline (prefix/numeric aware)
on isDGCompliant(verStr)
    if verStr is "" or verStr is "Unknown" then return false
    try
        if verStr starts with DG_APPROVED_VERSION then return true
        set lowest to do shell script "printf '%s\\n%s\\n' " & quoted form of DG_APPROVED_VERSION & " " & quoted form of verStr & " | sort -V | head -1"
        if lowest is DG_APPROVED_VERSION then return true
    end try
    return false
end isDGCompliant

-- Return list of visible, user-facing app names (excludes background, Finder, self)
on getOpenApps()
    set openList to {}
    try
        tell application "System Events"
            set openList to name of (every process whose background only is false and visible is true)
        end tell
    end try
    set filtered to {}
    repeat with a in openList
        set nm to (a as text)
        if nm is not "Finder" and nm is not appTitle and nm does not contain "TAHOE-DG" then
            set end of filtered to nm
        end if
    end repeat
    return filtered
end getOpenApps

-- Big warning + Quit All loop; returns true when clear, false if user cancels
on ensureAppsClosed()
    repeat
        set apps to my getOpenApps()
        if (count of apps) is 0 then
            my logLine("Open-apps gate: all clear.")
            return true
        end if
        set appText to ""
        repeat with a in apps
            set appText to appText & "   • " & (a as text) & return
        end repeat
        my logLine("Open-apps gate: still open -> " & (appText))
        set warnMsg to "⚠️  ALL OPEN APPS MUST BE CLOSED  ⚠️" & return & return & "This update will RESTART your Mac. To prevent any loss of unsaved work, the following applications must be closed first:" & return & return & appText & return & "Choose ‘Quit All Apps’ and save your work when prompted, or close them manually and re-check."
        try
            set r to display dialog warnMsg buttons {"Cancel", "Re-Check", "Quit All Apps"} default button "Quit All Apps" cancel button "Cancel" with title (appTitle & " — Close Apps Required") with icon caution
        on error number -128
            return false
        end try
        set act to button returned of r
        if act is "Quit All Apps" then
            repeat with a in apps
                set nm to (a as text)
                try
                    tell application nm to quit
                on error
                    try
                        tell application "System Events" to tell process nm to keystroke "q" using command down
                    end try
                end try
            end repeat
            delay 3 -- give apps time to present save prompts / close
        end if
        -- loop re-checks after Quit All or Re-Check
    end repeat
end ensureAppsClosed

-- Masked GUI password dialog + dscl validation + 3 retries
on collectCredential()
    set adminUser to consoleUser
    set a to 0
    repeat while a < 3
        try
            set r to display dialog "Enter the login password for " & adminUser & " to authorize these updates." & return & return & "Dots appear as you type so you can confirm your entry." default answer "" with title (appTitle & " — Administrator Password") buttons {"Cancel", "OK"} default button "OK" cancel button "Cancel" with icon note with hidden answer
        on error number -128
            return false
        end try
        set pw to text returned of r
        set okAuth to true
        try
            do shell script "/usr/bin/dscl /Local/Default -authonly " & quoted form of adminUser & " " & quoted form of pw
        on error
            set okAuth to false
        end try
        if okAuth then
            set adminPass to pw
            return true
        end if
        set a to a + 1
        if a < 3 then
            display dialog "That password was incorrect. Attempts remaining: " & (3 - a) buttons {"Try Again"} default button "Try Again" with icon caution with title appTitle
        end if
    end repeat
    return false
end collectCredential

on privSh(cmd)
    return do shell script cmd user name adminUser password adminPass with administrator privileges
end privSh

on runDigitalGuardian()
    set jamfPath to ""
    try
        set jamfPath to do shell script "command -v jamf || echo ''"
    end try
    if jamfPath is not "" then
        try
            my logLine("Triggering Jamf policy event: " & DG_POLICY_EVENT)
            set out to my privSh("/usr/local/bin/jamf policy -event " & quoted form of DG_POLICY_EVENT & " 2>&1")
            my logLine("jamf output: " & out)
            if out contains "No policies" then
                my logLine("No matching DG policy for that event — opening Self Service as fallback.")
                my openSelfService()
                return "Self Service opened (no policy matched event)"
            end if
            return "DG policy triggered via Jamf"
        on error errMsg
            my logLine("jamf policy error: " & errMsg & " — opening Self Service.")
            my openSelfService()
            return "Self Service opened (jamf error)"
        end try
    else
        my logLine("jamf binary not found — opening Self Service.")
        my openSelfService()
        return "Self Service opened (no jamf binary)"
    end if
end runDigitalGuardian

on openSelfService()
    try
        do shell script "open -a 'Self Service' 2>/dev/null || open -a 'Jamf Self Service' 2>/dev/null || true"
    end try
    try
        display dialog "Self Service has opened." & return & return & "Please locate ‘Digital Guardian " & DG_APPROVED_VERSION & ".x’ and click Install, then choose ‘Restart Now’ when prompted." buttons {"OK"} default button "OK" with title appTitle with icon note
    end try
end openSelfService

on listMacUpdates()
    set res to "none"
    try
        set res to my privSh("/usr/sbin/softwareupdate -l 2>&1")
        my logLine("softwareupdate -l:" & return & res)
    on error errMsg
        my logLine("softwareupdate -l error: " & errMsg)
    end try
    if res contains "No new software available" then
        return "No updates available"
    else
        return "Updates available (see log)"
    end if
end listMacUpdates

on installMacUpdates()
    try
        my logLine("Installing macOS updates (staged, restart deferred to end)...")
        set cmd to "/bin/sh -c 'printf %s " & quoted form of adminPass & " | /usr/sbin/softwareupdate -ia --agreeToLicense --user " & quoted form of adminUser & " --stdinpass 2>&1'"
        set out to my privSh(cmd)
        my logLine("softwareupdate -ia output:" & return & out)
        if out contains "No new software available" then
            return "No updates available"
        else
            return "macOS updates installed (restart to finish)"
        end if
    on error errMsg
        my logLine("softwareupdate install error: " & errMsg)
        return "Install error (see log)"
    end try
end installMacUpdates

on logLine(msg)
    set stamp to do shell script "date '+%H:%M:%S'"
    try
        do shell script "printf '%s  %s\\n' " & quoted form of stamp & " " & quoted form of msg & " >> " & quoted form of logPath
    end try
end logLine

on bailOut(msg)
    my logLine("ABORT: " & msg)
    display dialog msg buttons {"OK"} default button "OK" with title appTitle with icon stop
    error number -128
end bailOut

on writeReceipt(doMac, doDG, updatesFound, macResult, dgResult, dgAfter)
    set nowH to do shell script "date '+%b %d, %Y %I:%M %p %Z'"
    set macStatePill to my compliancePill(macCompliant)
    set dgStatePill to my compliancePill(dgCompliant)
    set macRow to my kpiRow("macOS Software Update", macResult)
    set dgRow to my kpiRow("Digital Guardian Update", dgResult)
    set html to "<!DOCTYPE html><html lang='en'><head><meta charset='UTF-8'><meta name='viewport' content='width=device-width, initial-scale=1'><title>TAHOE-DG Receipt - " & serialNum & "</title><style>:root{--gold:#d4af37;}*{box-sizing:border-box;font-family:-apple-system,Segoe UI,Roboto,sans-serif;}body{margin:0;background:#0a0c12;color:#e7ebf3;}.bg{position:fixed;inset:0;background:radial-gradient(1200px 600px at 20% -10%,rgba(212,175,55,.10),transparent),radial-gradient(900px 500px at 110% 20%,rgba(80,140,255,.08),transparent);}.wrap{position:relative;max-width:1000px;margin:0 auto;padding:36px 24px;}.hero{font-size:26px;font-weight:800;}.hero .c{color:var(--gold);}.sub{color:#9aa4b8;margin-top:6px;font-size:13px;}.card{background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.08);backdrop-filter:blur(14px);border-radius:16px;padding:22px;margin-top:22px;box-shadow:0 8px 30px rgba(0,0,0,.35);}.kpis{display:flex;gap:16px;flex-wrap:wrap;margin-top:20px;}.kpi{flex:1;min-width:160px;background:rgba(255,255,255,.03);border:1px solid rgba(255,255,255,.08);border-radius:14px;padding:16px;}.kpi .l{font-size:12px;color:#9aa4b8;text-transform:uppercase;letter-spacing:1px;}.kpi .v{font-size:18px;font-weight:800;margin-top:6px;word-break:break-word;}table{width:100%;border-collapse:collapse;margin-top:10px;font-size:14px;}th,td{text-align:left;padding:11px 12px;border-bottom:1px solid rgba(255,255,255,.07);}th{color:var(--gold);font-size:12px;text-transform:uppercase;letter-spacing:1px;}.pill{padding:3px 10px;border-radius:999px;font-size:12px;font-weight:700;}.pill.ok{background:rgba(57,217,138,.15);color:#39d98a;border:1px solid rgba(57,217,138,.4);}.pill.warn{background:rgba(212,175,55,.15);color:var(--gold);border:1px solid rgba(212,175,55,.4);}.foot{color:#6b7686;font-size:11px;margin-top:24px;text-align:center;}</style></head><body><div class='bg'></div><div class='wrap'><div class='hero'>TAHOE-<span class='c'>DG</span> — Update Receipt</div><div class='sub'>" & appSubtitle & " &bull; Host " & hostName & " &bull; Serial " & serialNum & " &bull; User " & adminUser & " &bull; " & nowH & "</div><div class='kpis'><div class='kpi'><div class='l'>macOS Before</div><div class='v'>" & osBefore & " (" & osBuild & ") " & macStatePill & "</div></div><div class='kpi'><div class='l'>Digital Guardian Detected</div><div class='v'>" & dgBefore & " " & dgStatePill & "</div></div><div class='kpi'><div class='l'>Digital Guardian After</div><div class='v'>" & dgAfter & "</div></div><div class='kpi'><div class='l'>Updates Found</div><div class='v'>" & updatesFound & "</div></div></div><div class='card'><h3 style='margin:0 0 6px;color:var(--gold);'>Task Results</h3><table><tr><th>Task</th><th>Result</th></tr>" & macRow & dgRow & "</table></div><div class='card'><h3 style='margin:0 0 6px;color:var(--gold);'>Compliance Baseline</h3><p style='color:#c3cad8;font-size:14px;'>macOS baseline: <b>Tahoe " & MACOS_MIN_MAJOR & "+</b>. Digital Guardian baseline: <b>" & DG_APPROVED_VERSION & "+</b>. Versions are checked BEFORE any app scan, credential prompt, or update; items already meeting baseline are skipped automatically. Full command output and the detected DG version are captured in <b>update.log</b>.</p></div><div class='foot'>Confidential — For internal use only. " & appTitle & " v" & appVersion & " &bull; Log: update.log</div></div></body></html>"
    try
        set fh to open for access (POSIX file htmlPath) with write permission
        set eof fh to 0
        write html to fh as «class utf8»
        close access fh
    on error
        try
            close access (POSIX file htmlPath)
        end try
    end try
end writeReceipt

on compliancePill(isOK)
    if isOK then
        return "<span class='pill ok'>OK</span>"
    else
        return "<span class='pill warn'>UPDATE</span>"
    end if
end compliancePill

on kpiRow(label, val)
    set cls to "warn"
    if (val contains "installed") or (val contains "triggered") or (val contains "No updates") or (val contains "up to date") or (val contains "already current") then set cls to "ok"
    return "<tr><td>" & label & "</td><td><span class='pill " & cls & "'>" & val & "</span></td></tr>"
end kpiRow
