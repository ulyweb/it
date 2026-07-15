# The One-Liner (empties the Dock's app side)


## 🧹 Want to clear the right side too (folders/stacks like Downloads)?

Chain both arrays in a single command:

````
Shelldefaults write com.apple.dock persistent-apps -array; defaults write com.apple.dock persistent-others -array; killall DockShow more lines
````

The persistent-others key controls the folders/stacks section on the right of the divider.


### 🎨 Script Editor version (if you prefer a clickable app)


Open Script Editor → new document →:

````
AppleScriptdo shell script "defaults write com.apple.dock persistent-apps -array; defaults write com.apple.dock persistent-others -array; killall Dock"
````

Then File → Export → File Format: Application to save it as a double-click "Clear Dock" app you can reuse across your executive fleet.

#### 💡 Tip: If you're deploying this to executive Macs, run the empty-array version (not defaults delete) so you get a truly blank slate, then push your approved app set afterward — that gives you a consistent, standardized Dock layout across devices.



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
* **Self Service (Jamf)** → normally `/Applications/Self Service.app` — but some deployments name it `Self Service+.app`
* **GlobalProtect** → usually `/Applications/GlobalProtect.app` (occasionally lives under a vendor subfolder)
* **Personal Print Manager (LRS)** → confirm the exact bundle name; sometimes it's `Personal Print Manager.app` and sometimes shortened
* **Google Chrome** → `Google Chrome.app` (with the space)
* **Safari** → `/Applications/Safari.app` ✅ (standard location)
* **Activity Monitor** → `/System/Applications/Utilities/Activity Monitor.app` — Apple moved system utilities into the read-only System volume, so `/Applications/Utilities/...` will **fail** on modern macOS. The path above is correct for Tahoe 26.


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

