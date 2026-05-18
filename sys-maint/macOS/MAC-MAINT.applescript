-- ============================================================
--  macOS Maintenance Dashboard — Native Launcher
--  HOW TO USE:
--    1. Open this file in Script Editor.app
--    2. Click the Compile button (hammer icon) to verify
--    3. File → Export → File Format: Application
--    4. Save as "MAC-MAINT.app" in the same folder as
--       launcher_mac.py and dashboard_mac.html
--    5. Double-click MAC-MAINT.app to launch
-- ============================================================

property appName : "macOS Maintenance Dashboard"
property appPort : "9292"
property pythonScript : "launcher_mac.py"
property htmlFile : "dashboard_mac.html"
property brewInstallURL : "https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh"

-- ── Entry point ────────────────────────────────────────────
on run
	
	-- Get the folder this .app lives in
	set appPath to POSIX path of (path to me)
	set scriptFolder to do shell script "dirname " & quoted form of appPath
	
	set launcherPath to scriptFolder & "/" & pythonScript
	set htmlPath to scriptFolder & "/" & htmlFile
	
	-- ── Check required files ──────────────────────────────────
	set launcherExists to do shell script "[ -f " & quoted form of launcherPath & " ] && echo yes || echo no"
	if launcherExists is "no" then
		display alert "Missing File" message "launcher_mac.py not found in:" & return & scriptFolder & return & return & "Put all 3 files in the same folder as this app." as critical
		return
	end if
	
	set htmlExists to do shell script "[ -f " & quoted form of htmlPath & " ] && echo yes || echo no"
	if htmlExists is "no" then
		display alert "Missing File" message "dashboard_mac.html not found in:" & return & scriptFolder & return & return & "Put all 3 files in the same folder as this app." as critical
		return
	end if
	
	-- ── Step 1: Find or install Python 3 ─────────────────────
	display notification "Checking for Python 3..." with title appName
	
	set pythonBin to findPython()
	
	if pythonBin is "" then
		display notification "Python 3 not found — installing now..." with title appName
		set pythonBin to installPython()
		if pythonBin is "" then
			display alert "Setup Failed" message "Python 3 could not be installed automatically." & return & return & "Manual option: https://www.python.org/downloads/macos/" as critical
			return
		end if
	end if
	
	-- ── Step 2: Check if port is already in use ──────────────
	set portPID to do shell script "lsof -ti tcp:" & appPort & " 2>/dev/null | head -1 || echo ''"
	
	if portPID is not "" then
		set portChoice to button returned of (display alert "Port " & appPort & " In Use" message "Something is already running on port " & appPort & "." & return & return & "The dashboard may already be open." buttons {"Kill & Restart", "Open Dashboard", "Cancel"} default button "Open Dashboard" as warning)
		
		if portChoice is "Cancel" then
			return
		else if portChoice is "Open Dashboard" then
			do shell script "open http://localhost:" & appPort
			return
		else
			do shell script "lsof -ti tcp:" & appPort & " | xargs kill -9 2>/dev/null; sleep 1" with administrator privileges
		end if
	end if
	
	-- ── Step 3: Start the Python server ──────────────────────
	display notification "Starting server on port " & appPort & "..." with title appName
	
	set launchCmd to "cd " & quoted form of scriptFolder & " && " & quoted form of pythonBin & " " & quoted form of launcherPath & " > /tmp/mac-maint.log 2>&1 &"
	do shell script launchCmd
	
	-- Wait up to 5 seconds for the server to come up
	set serverReady to false
	repeat 5 times
		delay 1
		set checkResult to do shell script "lsof -ti tcp:" & appPort & " 2>/dev/null | head -1 || echo ''"
		if checkResult is not "" then
			set serverReady to true
			exit repeat
		end if
	end repeat
	
	if serverReady is false then
		set logOutput to do shell script "cat /tmp/mac-maint.log 2>/dev/null | tail -5 || echo 'No log output'"
		display alert "Server Failed to Start" message "The server did not come up on port " & appPort & "." & return & return & "Last log:" & return & logOutput as critical
		return
	end if
	
	-- Open browser
	do shell script "open http://localhost:" & appPort
	
	-- ── Step 4: Running — offer stop control ─────────────────
	set runChoice to button returned of (display alert appName & " is Running" message "Dashboard is live at:" & return & "http://localhost:" & appPort & return & return & "Use the Shutdown button inside the dashboard, or click Stop Server below." buttons {"Stop Server", "Keep Running"} default button "Keep Running")
	
	if runChoice is "Stop Server" then
		stopServer()
	end if
	
end run

-- ── Find Python 3 in common locations ─────────────────────
on findPython()
	
	set loc1 to "/opt/homebrew/bin/python3"
	set loc2 to "/usr/local/bin/python3"
	set loc3 to "/usr/bin/python3"
	set loc4 to "/opt/homebrew/bin/python3.12"
	set loc5 to "/opt/homebrew/bin/python3.11"
	
	try
		set r to do shell script loc1 & " --version 2>&1"
		if r contains "Python 3" then return loc1
	end try
	
	try
		set r to do shell script loc2 & " --version 2>&1"
		if r contains "Python 3" then return loc2
	end try
	
	try
		set r to do shell script loc3 & " --version 2>&1"
		if r contains "Python 3" then return loc3
	end try
	
	try
		set r to do shell script loc4 & " --version 2>&1"
		if r contains "Python 3" then return loc4
	end try
	
	try
		set r to do shell script loc5 & " --version 2>&1"
		if r contains "Python 3" then return loc5
	end try
	
	-- Last resort: ask the login shell
	try
		set r to do shell script "python3 --version 2>&1"
		if r contains "Python 3" then return "python3"
	end try
	
	return ""
	
end findPython

-- ── Install Homebrew then Python 3 ────────────────────────
on installPython()
	
	-- Check if Homebrew is already present
	set brewBin to ""
	
	set appleChip to do shell script "[ -f /opt/homebrew/bin/brew ] && echo yes || echo no"
	set intelChip to do shell script "[ -f /usr/local/bin/brew ] && echo yes || echo no"
	
	if appleChip is "yes" then
		set brewBin to "/opt/homebrew/bin/brew"
	else if intelChip is "yes" then
		set brewBin to "/usr/local/bin/brew"
	end if
	
	-- Install Homebrew if not found
	if brewBin is "" then
		
		display notification "Installing Homebrew (1-3 min)..." with title appName subtitle "Please wait — do not close this"
		
		set brewCmd to "NONINTERACTIVE=1 /bin/bash -c \"$(curl -fsSL " & brewInstallURL & ")\" >> /tmp/mac-maint-brew.log 2>&1"
		
		try
			do shell script brewCmd with administrator privileges
		on error brewErr
			display alert "Homebrew Install Failed" message brewErr & return & return & "See /tmp/mac-maint-brew.log for details." as critical
			return ""
		end try
		
		-- Re-detect location after install
		set appleChip to do shell script "[ -f /opt/homebrew/bin/brew ] && echo yes || echo no"
		set intelChip to do shell script "[ -f /usr/local/bin/brew ] && echo yes || echo no"
		
		if appleChip is "yes" then
			set brewBin to "/opt/homebrew/bin/brew"
		else if intelChip is "yes" then
			set brewBin to "/usr/local/bin/brew"
		else
			display alert "Homebrew Not Found After Install" message "Homebrew installed but could not be located." & return & "See /tmp/mac-maint-brew.log for details." as critical
			return ""
		end if
		
		display notification "Homebrew installed successfully." with title appName
		
	end if
	
	-- Install Python 3 via Homebrew
	display notification "Installing Python 3 via Homebrew..." with title appName subtitle "Almost done..."
	
	try
		do shell script brewBin & " install python3 >> /tmp/mac-maint-brew.log 2>&1"
	on error pyErr
		display alert "Python 3 Install Failed" message pyErr & return & return & "See /tmp/mac-maint-brew.log for details." as critical
		return ""
	end try
	
	-- Verify Python 3 is now reachable
	set pythonBin to findPython()
	
	if pythonBin is not "" then
		display notification "Python 3 ready." with title appName
		return pythonBin
	else
		display alert "Python 3 Not Detected" message "Homebrew finished but python3 was still not found in expected paths." & return & return & "Open Terminal and run: python3 --version" as warning
		return ""
	end if
	
end installPython

-- ── Stop the server ────────────────────────────────────────
on stopServer()
	try
		do shell script "lsof -ti tcp:" & appPort & " | xargs kill -9 2>/dev/null; echo done"
		display notification "Server stopped. Port " & appPort & " is now free." with title appName
	on error
		display notification "Server was already stopped." with title appName
	end try
end stopServer
