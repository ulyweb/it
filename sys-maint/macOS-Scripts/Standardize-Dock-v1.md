## 💡 Recommended Professional Enhancements

* **Pre-flight path validation** — check each app exists *before* pinning, skip missing ones gracefully, and report them so you never get silent skips or `?` placeholders. ✅
* **Automatic backup + rollback** — snapshot the current `com.apple.dock.plist` (timestamped) before any change, so any tech can instantly restore the prior Dock if needed. 🔄
* **Confirmation gate on destructive actions** — a "are you sure?" prompt before the factory reset to prevent accidental wipes. ⚠️
* **Logging** — append every run to a log file (`/Users/Shared/Standardize-Dock/logs/`) with timestamp, user, action, and results for audit/metrics. 📋
* **Summary report** — end-of-run dialog showing what was applied, skipped, and where the backup/log live. 📊
* **Error handling** — `try` blocks so one bad path never halts the whole run.
* **Version metadata header** — for change tracking across your fleet.

## 🎛️ Production Script — "Standardize Dock v1.0"

Open **Script Editor** → paste → **File → Export → File Format: Application** (save as `Standardize-Dock` in `/Users/Shared`):

```applescript
-- ============================================================
--  Standardize Dock
--  Version : 1.0
--  Author  : ulyweb
--  Date    : 2026-07-15
--  Purpose : Apply a standard core-app Dock, factory-reset the
--            Dock, or restore a previous Dock from backup.
--            Includes path validation, auto-backup, logging,
--            confirmation gates, and a summary report.
-- ============================================================

-- ---------- Configuration ----------
set appList to {"/Applications/Safari.app", "/Applications/Google Chrome.app", "/Applications/zoom.us.app", "/Applications/Box.app", "/Applications/Microsoft Outlook.app", "/Applications/Microsoft Word.app", "/Applications/Microsoft Excel.app", "/Applications/Microsoft PowerPoint.app", "/Applications/Self Service.app", "/Applications/Personal Print Manager.app", "/Applications/GlobalProtect.app", "/System/Applications/Utilities/Activity Monitor.app"}

set baseDir to "/Users/Shared/Standardize-Dock"
set backupDir to baseDir & "/backups"
set logDir to baseDir & "/logs"
set plistPath to (POSIX path of (path to home folder)) & "Library/Preferences/com.apple.dock.plist"

-- ---------- Setup working folders ----------
do shell script "mkdir -p " & quoted form of backupDir & " " & quoted form of logDir

set timeStamp to do shell script "date +%Y%m%d_%H%M%S"
set currentUser to do shell script "whoami"
set logFile to logDir & "/standardize-dock.log"

-- ---------- Helper: write to log ----------
on writeLog(msg, logFile, currentUser)
	set ts to do shell script "date '+%Y-%m-%d %H:%M:%S'"
	do shell script "echo " & quoted form of (ts & " [" & currentUser & "] " & msg) & " >> " & quoted form of logFile
end writeLog

-- ---------- Main menu ----------
set actionChoice to choose from list {"Standard Core Apps Dock", "Fully Reset Dock to Factory Defaults", "Restore Previous Dock from Backup"} with title "Standardize Dock v1.0" with prompt "Select the Dock action to perform:" default items {"Standard Core Apps Dock"} without empty selection allowed

if actionChoice is false then
	my writeLog("Action cancelled at main menu.", logFile, currentUser)
	return
end if
set choice to item 1 of actionChoice

-- ============================================================
--  OPTION 1: Standard Core Apps Dock
-- ============================================================
if choice is "Standard Core Apps Dock" then
	
	-- Pre-flight: validate app paths
	set foundApps to {}
	set missingApps to {}
	repeat with a in appList
		set aPath to contents of a
		if (do shell script "[ -e " & quoted form of aPath & " ] && echo yes || echo no") is "yes" then
			set end of foundApps to aPath
		else
			set end of missingApps to aPath
		end if
	end repeat
	
	-- Report missing before proceeding
	if (count of missingApps) > 0 then
		set missingText to ""
		repeat with m in missingApps
			set missingText to missingText & "• " & (contents of m) & return
		end repeat
		set proceed to button returned of (display dialog "The following apps were NOT found and will be skipped:" & return & return & missingText & return & "Proceed with the " & (count of foundApps) & " apps that were found?" buttons {"Cancel", "Proceed"} default button "Proceed" with title "Standardize Dock — Path Check" with icon caution)
		if proceed is "Cancel" then
			my writeLog("Cancelled after path check. Missing: " & (count of missingApps), logFile, currentUser)
			return
		end if
	end if
	
	-- Backup current Dock plist
	try
		do shell script "cp " & quoted form of plistPath & " " & quoted form of (backupDir & "/com.apple.dock_" & timeStamp & ".plist")
	end try
	
	-- Apply layout
	do shell script "defaults write com.apple.dock persistent-apps -array"
	set appliedCount to 0
	repeat with a in foundApps
		try
			do shell script "defaults write com.apple.dock persistent-apps -array-add '<dict><key>tile-data</key><dict><key>file-data</key><dict><key>_CFURLString</key><string>" & (contents of a) & "</string><key>_CFURLStringType</key><integer>0</integer></dict></dict></dict>'"
			set appliedCount to appliedCount + 1
		end try
	end repeat
	do shell script "killall Dock"
	
	my writeLog("Standard Core Apps applied. Pinned: " & appliedCount & " | Skipped: " & (count of missingApps), logFile, currentUser)
	
	display dialog "✅ Standard Core Apps Dock applied." & return & return & "Apps pinned: " & appliedCount & return & "Apps skipped: " & (count of missingApps) & return & "Backup: " & backupDir & return & "Log: " & logFile buttons {"OK"} default button "OK" with title "Standardize Dock — Complete"
	
-- ============================================================
--  OPTION 2: Factory Reset (with confirmation)
-- ============================================================
else if choice is "Fully Reset Dock to Factory Defaults" then
	
	set confirmReset to button returned of (display dialog "⚠️ This will DELETE all custom Dock settings (pinned apps, size, position, magnification) and restore macOS factory defaults." & return & return & "A backup will be saved first. Continue?" buttons {"Cancel", "Reset Dock"} default button "Cancel" with title "Confirm Factory Reset" with icon caution)
	if confirmReset is "Cancel" then
		my writeLog("Factory reset cancelled by user.", logFile, currentUser)
		return
	end if
	
	-- Backup then reset
	try
		do shell script "cp " & quoted form of plistPath & " " & quoted form of (backupDir & "/com.apple.dock_" & timeStamp & ".plist")
	end try
	do shell script "defaults delete com.apple.dock; killall Dock"
	
	my writeLog("Dock factory reset performed. Backup: com.apple.dock_" & timeStamp & ".plist", logFile, currentUser)
	
	display dialog "✅ Dock reset to macOS factory defaults." & return & return & "Backup saved to:" & return & backupDir buttons {"OK"} default button "OK" with title "Standardize Dock — Complete"
	
-- ============================================================
--  OPTION 3: Restore Previous Dock from Backup
-- ============================================================
else if choice is "Restore Previous Dock from Backup" then
	
	set backupFiles to paragraphs of (do shell script "ls -1t " & quoted form of backupDir & " 2>/dev/null || true")
	if backupFiles is {} or (item 1 of backupFiles) is "" then
		display dialog "No backups were found in:" & return & backupDir buttons {"OK"} default button "OK" with title "Standardize Dock — Restore" with icon caution
		my writeLog("Restore attempted but no backups found.", logFile, currentUser)
		return
	end if
	
	set pickBackup to choose from list backupFiles with title "Restore Previous Dock" with prompt "Select a backup to restore (newest first):" without empty selection allowed
	if pickBackup is false then
		my writeLog("Restore cancelled at backup selection.", logFile, currentUser)
		return
	end if
	
	set selectedBackup to item 1 of pickBackup
	do shell script "cp " & quoted form of (backupDir & "/" & selectedBackup) & " " & quoted form of plistPath & "; killall Dock"
	
	my writeLog("Dock restored from backup: " & selectedBackup, logFile, currentUser)
	
	display dialog "✅ Dock restored from backup:" & return & return & selectedBackup buttons {"OK"} default button "OK" with title "Standardize Dock — Complete"
	
end if
```

## 📝 What This Version Gives You

| Feature               | Benefit                                                                            |
| --------------------- | ---------------------------------------------------------------------------------- |
| **Path validation**   | Reports missing apps *before* applying; no silent skips or `?` icons               |
| **Auto-backup**       | Every change snapshots the Dock plist to `/Users/Shared/Standardize-Dock/backups/` |
| **Restore option**    | Third menu item lets any tech roll back to a prior Dock (newest-first list)        |
| **Confirmation gate** | Factory reset requires explicit "Reset Dock" click — no accidents                  |
| **Logging**           | Append-only audit log with timestamp, user, action, and counts                     |
| **Summary dialog**    | Clear end-of-run report: applied / skipped / backup / log paths                    |
| **Error handling**    | `try` blocks keep one bad path from halting the run                                |

The factory-reset path still uses the standard `defaults delete com.apple.dock; killall Dock` so macOS regenerates its original layout. [\[idownloadblog.com\]](https://www.idownloadblog.com/2023/05/23/how-to-reset-mac-dock/), [\[discussion....apple.com\]](https://discussions.apple.com/thread/254797028)


