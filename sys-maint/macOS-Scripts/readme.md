# The One-Liner (empties the Dock's app side)


## 🧹 Want to clear the right side too (folders/stacks like Downloads)?

Chain both arrays in a single command:

````shell
defaults write com.apple.dock persistent-apps -array; defaults write com.apple.dock persistent-others -array; killall DockShow more lines
````

The persistent-others key controls the folders/stacks section on the right of the divider.

### To fully reset the Dock to factory defaults instead (size, magnification, position, and re-populate Apple's default apps), use:

```
defaults delete com.apple.dock && killall Dock
```

### 🎨 Script Editor version (if you prefer a clickable app)


Open Script Editor → new document →:

````AppleScript
do shell script "defaults write com.apple.dock persistent-apps -array; defaults write com.apple.dock persistent-others -array; killall Dock"
````

Then File → Export → File Format: Application to save it as a double-click "Clear Dock" app you can reuse across your executive fleet.



---


## 📝 Verify Your App Paths First

App bundle names must match **exactly** or the icon won't pin. Quick check:


## 📝 Verify Paths First
```bash
ls -d /Applications/zoom.us.app /Applications/Box.app "/Applications/Microsoft Outlook.app" "/Applications/Microsoft Word.app" "/Applications/Microsoft Excel.app" "/Applications/Microsoft PowerPoint.app" "/Applications/Safari.app" "/System/Applications/Utilities/Activity Monitor.app" "/Applications/Google Chrome.app" "/Applications/Self Service.app" "/Applications/Personal Print Manager.app" "/Applications/GlobalProtect.app"
```


Key notes:
* **Zoom** → `zoom.us.app` (lowercase, not "Zoom.app")
* **Box Drive** → usually `Box.app` (older installs may be `Box Drive.app`)
* **Microsoft apps** → full name with space, e.g. `Microsoft Outlook.app`
* **Self Service** → normally `/Applications/Self Service.app` — but some deployments name it `Self Service+.app`
* **GlobalProtect** → usually `/Applications/GlobalProtect.app` (occasionally lives under a vendor subfolder)
* **Personal Print Manager** → confirm the exact bundle name; sometimes it's `Personal Print Manager.app` and sometimes shortened
* **Google Chrome** → `Google Chrome.app` (with the space)
* **Safari** → `/Applications/Safari.app` ✅ (standard location)
* **Activity Monitor** → `/System/Applications/Utilities/Activity Monitor.app` — Apple moved system utilities into the read-only System volume.


## ✅ One-Liner — Full 12-App Set (native `defaults`, clear + rebuild)

```bash
defaults write com.apple.dock persistent-apps -array; for app in "/Applications/Safari.app" "/Applications/Google Chrome.app" "/Applications/zoom.us.app" "/Applications/Box.app" "/Applications/Microsoft Outlook.app" "/Applications/Microsoft Word.app" "/Applications/Microsoft Excel.app" "/Applications/Microsoft PowerPoint.app" "/Applications/Self Service.app" "/Applications/Personal Print Manager.app" "/Applications/GlobalProtect.app" "/System/Applications/Utilities/Activity Monitor.app"; do defaults write com.apple.dock persistent-apps -array-add "<dict><key>tile-data</key><dict><key>file-data</key><dict><key>_CFURLString</key><string>$app</string><key>_CFURLStringType</key><integer>0</integer></dict></dict></dict>"; done; killall Dock
```

## 🎨 Script Editor Version (clickable "Standardize Dock" app)

Open **Script Editor** → new document → paste → **File → Export → File Format: Application**:

```applescript
set appList to {"/Applications/Safari.app", "/Applications/Google Chrome.app", "/Applications/zoom.us.app", "/Applications/Box.app", "/Applications/Microsoft Outlook.app", "/Applications/Microsoft Word.app", "/Applications/Microsoft Excel.app", "/Applications/Microsoft PowerPoint.app", "/Applications/Self Service.app", "/Applications/Personal Print Manager.app", "/Applications/GlobalProtect.app", "/System/Applications/Utilities/Activity Monitor.app"}
do shell script "defaults write com.apple.dock persistent-apps -array"
repeat with a in appList
	do shell script "defaults write com.apple.dock persistent-apps -array-add '<dict><key>tile-data</key><dict><key>file-data</key><dict><key>_CFURLString</key><string>" & a & "</string><key>_CFURLStringType</key><integer>0</integer></dict></dict></dict>'"
end repeat
do shell script "killall Dock"
```


---


Here's a **Script Editor app with a selection menu** that lets you (or any tech) pick between the two actions at runtime: apply the **full 12-app standardized set**, or **fully reset the Dock to factory defaults**. 🎯

## 🎛️ Script Editor Version — Interactive Chooser

Open **Script Editor** → new document → paste the script below → **File → Export → File Format: Application** (save as `Standardize-Dock` in `/Users/Shared`):

```applescript
-- ============================================================
--  Standardize Dock — Interactive Chooser
--  Option 1: Apply full 12-app standardized set
--  Option 2: Fully reset Dock to macOS factory defaults
-- ============================================================

set appList to {"/Applications/Safari.app", "/Applications/Google Chrome.app", "/Applications/zoom.us.app", "/Applications/Box.app", "/Applications/Microsoft Outlook.app", "/Applications/Microsoft Word.app", "/Applications/Microsoft Excel.app", "/Applications/Microsoft PowerPoint.app", "/Applications/Self Service.app", "/Applications/Personal Print Manager.app", "/Applications/GlobalProtect.app", "/System/Applications/Utilities/Activity Monitor.app"}

set actionChoice to choose from list {"Core Standard Dock App", "Fully Reset Dock to Factory Defaults"} with title "Standardize Dock" with prompt "Select the Dock action to perform:" default items {"Core Standard Dock App"} without empty selection allowed

if actionChoice is false then
	-- User clicked Cancel
	return
end if

set choice to item 1 of actionChoice

if choice is "Core Standard Dock App" then
	-- Clear the app side, then pin the approved 12-app set
	do shell script "defaults write com.apple.dock persistent-apps -array"
	repeat with a in appList
		do shell script "defaults write com.apple.dock persistent-apps -array-add '<dict><key>tile-data</key><dict><key>file-data</key><dict><key>_CFURLString</key><string>" & a & "</string><key>_CFURLStringType</key><integer>0</integer></dict></dict></dict>'"
	end repeat
	do shell script "killall Dock"
	display notification "Core Standard Dock App." with title "Core Standard Dock App" subtitle "Complete"
	
else if choice is "Fully Reset Mac's Dock to Factory Defaults" then
	-- Wipe all custom Dock settings; macOS regenerates factory defaults
	do shell script "defaults delete com.apple.dock; killall Dock"
	display notification "Dock reset to macOS factory defaults." with title "Standardize Dock" subtitle "Complete"
end if
```

## 📝 How It Works

* On launch it shows a **`choose from list`** dialog with two clear options and a confirmation notification when done. 🖱️
* **Option 1 – Apply Full 12-App Set:** empties the app side, then pins your approved 12 apps in the logical order (browsers → comms/storage → Office → IT utilities → system monitor). 
* **Option 2 – Factory Reset:** runs `defaults delete com.apple.dock; killall Dock`, which wipes all custom Dock settings and lets macOS regenerate the original factory layout, size, magnification, and position. 
* **Cancel** exits cleanly with no changes.


