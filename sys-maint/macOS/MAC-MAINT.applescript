-- ============================================================
--  macOS Maintenance Dashboard — Native Launcher
--  HOW TO USE:
--    1. Open this file in Script Editor.app
--    2. File → Export → File Format: Application
--    3. Save as "MAC-MAINT.app" in the same folder as
--       launcher_mac.py and dashboard_mac.html
--    4. Double-click MAC-MAINT.app to launch
-- ============================================================

-- ── Configuration ─────────────────────────────────────────
property appName : "macOS Maintenance Dashboard"
property appPort : "9292"
property pythonScript : "launcher_mac.py"
property htmlFile : "dashboard_mac.html"
property brewInstallURL : "https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh"

-- ── Entry point ────────────────────────────────────────────
on run
	
	-- Find the folder this .app lives in
	set appPath to POSIX path of (path to me)
	-- path to me returns the .app bundle path; get its parent folder
	set scriptFolder to do shell script "dirname " & quoted form of appPath
	
	-- ── Step 1: Verify required files exist ──────────────────
	set launcherPath to scriptFolder & "/" & pythonScript
	set htmlPath to scriptFolder & "/" & htmlFile
	
	set launcherExists to do shell script "[ -f " & quoted form of launcherPath & " ] && echo yes || echo no"
	set htmlExists to do shell script "[ -f " & quoted form of htmlPath & " ] && echo yes || echo no"
	
	if launcherExists is "no" then
		display alert "Missing File" message "launcher_mac.py was not found in:" & return & scriptFolder & return & return & "Make sure all 3 files are in the same folder as this app." as critical
		return
	end if
	
	if htmlExists is "no" then
		display alert "Missing File" message "dashboard_mac.html was not found in:" & return & scriptFolder & return & return & "Make sure all 3 files are in the same folder as this app." as critical
		return
	end if
	
	-- ── Step 2: Check for Python 3 ───────────────────────────
	showProgress("Checking for Python 3...")
	
	set pythonBin to findPython()
	
	if pythonBin is "" then
		-- Python not found — install it
		set pythonBin to installPython(scriptFolder)
		if pythonBin is "" then
			display alert "Installation Failed" message "Python 3 could not be installed automatically." & return & return & "Manual fallback: visit https://www.python.org/downloads/macos/" as critical
			return
		end if
	end if
	
	-- ── Step 3: Check if port is already in use ──────────────
	set portCheck to do shell script "lsof -ti tcp:" & appPort & " 2>/dev/null | head -1 || echo ''"
	if portCheck is not "" then
		set userChoice to button returned of (display alert "Port " & appPort & " Already in Use" message "The dashboard may already be running. A process is using port " & appPort & "." & return & return & "Kill it and start fresh, or open the existing dashboard?" buttons {"Kill & Restart", "Open Dashboard", "Cancel"} default button "Open Dashboard")
		if userChoice is "Cancel" then
			return
		else if userChoice is "Open Dashboard" then
			do shell script "open http://localhost:" & appPort
			return
		else
			-- Kill existing process on that port
			do shell script "lsof -ti tcp:" & appPort & " | xargs kill -9 2>/dev/null; sleep 1" with administrator privileges
		end if
	end if
	
	-- ── Step 4: Launch the Python server ─────────────────────
	showProgress("Starting server on port " & appPort & "...")
	
	-- Build the launch command — runs in background, opens browser
	set launchCmd to "cd " & quoted form of scriptFolder & " && " & quoted form of pythonBin & " " & quoted form of launcherPath & " &"
	
	do shell script launchCmd
	
	-- Give the server a moment to start
	delay 2
	
	-- Verify it started successfully
	set serverCheck to do shell script "lsof -ti tcp:" & appPort & " 2>/dev/null | head -1 || echo ''"
	if serverCheck is "" then
		display alert "Server Failed to Start" message "The Python server did not start on port " & appPort & "." & return & return & "Try running: python3 launcher_mac.py" & return & "in Terminal to see the error." as critical
		return
	end if
	
	-- Open the dashboard in the default browser
	do shell script "open http://localhost:" & appPort
	
	-- ── Step 5: Notify user and offer stop control ────────────
	set userChoice to button returned of (display alert appName & " is Running" message "Dashboard is live at:" & return & "http://localhost:" & appPort & return & return & "Use the Shutdown button inside the dashboard to stop it, or click Stop Server below." buttons {"Stop Server", "Keep Running"} default button "Keep Running")
	
	if userChoice is "Stop Server" then
		stopServer()
	end if
	
end run

-- ── Find Python 3 ─────────────────────────────────────────
on findPython()
	-- Check common locations: PATH, Homebrew Apple Silicon, Homebrew Intel, pyenv
	set candidates to {"/opt/homebrew/bin/python3", "/usr/local/bin/python3", "/usr/bin/python3", "/opt/homebrew/bin/python", "python3", "python"}
	
	repeat with candidate in candidates
		try
			set ver to do shell script candidate & " --version 2>&1"
			if ver contains "Python 3" then
				return candidate
			end if
		end try
	end repeat
	
	-- One more try: ask the shell with full PATH
	try
		set ver to do shell script "source ~/.zshrc 2>/dev/null; source ~/.bash_profile 2>/dev/null; python3 --version 2>&1"
		if ver contains "Python 3" then
			return "python3"
		end if
	end try
	
	return ""
end findPython

-- ── Install Python 3 (fully automated) ────────────────────
on installPython(scriptFolder)
	
	-- Check if Homebrew is available
	set brewBin to ""
	repeat with candidate in {"/opt/homebrew/bin/brew", "/usr/local/bin/brew"}
		try
			do shell script candidate & " --version 2>&1"
			set brewBin to candidate
			exit repeat
		end try
	end repeat
	
	if brewBin is "" then
		-- Homebrew not found — need to install it first
		-- Show native progress dialog
		display notification "Installing Homebrew — this takes 1-3 minutes..." with title appName subtitle "Please wait, do not close this dialog"
		
		set installBrewCmd to "NONINTERACTIVE=1 /bin/bash -c \"$(curl -fsSL " & brewInstallURL & ")\" 2>&1"
		
		try
			do shell script installBrewCmd with administrator privileges
		on error errMsg
			display alert "Homebrew Installation Failed" message errMsg as critical
			return ""
		end try
		
		-- Detect which prefix was used (Apple Silicon vs Intel)
		try
			do shell script "/opt/homebrew/bin/brew --version 2>&1"
			set brewBin to "/opt/homebrew/bin/brew"
		on try error
			try
				do shell script "/usr/local/bin/brew --version 2>&1"
				set brewBin to "/usr/local/bin/brew"
			on error
				display alert "Homebrew Not Found" message "Homebrew installed but could not be located. Try restarting and running again." as critical
				return ""
			end try
		end try
		
		display notification "Homebrew installed successfully." with title appName
	end if
	
	-- Now install Python 3 via Homebrew
	display notification "Installing Python 3 via Homebrew..." with title appName subtitle "Almost there..."
	
	try
		do shell script brewBin & " install python3 2>&1"
	on error errMsg
		display alert "Python 3 Installation Failed" message errMsg as critical
		return ""
	end try
	
	-- Verify Python 3 is now available
	set pythonBin to findPython()
	
	if pythonBin is not "" then
		display notification "Python 3 installed successfully!" with title appName
		return pythonBin
	else
		display alert "Python 3 Not Found After Install" message "Homebrew finished but python3 was still not found in expected paths." & return & return & "Try opening a new Terminal and running: python3 --version" as warning
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

-- ── Simple progress notification helper ───────────────────
on showProgress(msg)
	display notification msg with title appName
end showProgress
