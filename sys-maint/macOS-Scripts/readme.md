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


Yes! And since you prefer **native bash with zero dependencies** (no Homebrew/dockutil needed), here's a clean one-liner that pins all your core apps in one shot. 👇

## ✅ One-Liner — Add Your Core App Set (native `defaults`)

```bash
for app in "/Applications/zoom.us.app" "/Applications/Box.app" "/Applications/Microsoft Outlook.app" "/Applications/Microsoft Word.app" "/Applications/Microsoft Excel.app" "/Applications/Microsoft PowerPoint.app"; do defaults write com.apple.dock persistent-apps -array-add "<dict><key>tile-data</key><dict><key>file-data</key><dict><key>_CFURLString</key><string>$app</string><key>_CFURLStringType</key><integer>0</integer></dict></dict></dict>"; done; killall Dock
```

This loops through each app path, appends it to the Dock's `persistent-apps` array, then restarts the Dock so the icons appear immediately. [\[discussion....apple.com\]](https://discussions.apple.com/thread/254797028), [\[discussion....apple.com\]](https://discussions.apple.com/thread/254332755)

## 🔄 Best Practice: Clear + Rebuild in One Go

For a clean, **standardized executive Dock** (empty it first, then add only your approved apps), chain it:

```bash
defaults write com.apple.dock persistent-apps -array; for app in "/Applications/zoom.us.app" "/Applications/Box.app" "/Applications/Microsoft Outlook.app" "/Applications/Microsoft Word.app" "/Applications/Microsoft Excel.app" "/Applications/Microsoft PowerPoint.app"; do defaults write com.apple.dock persistent-apps -array-add "<dict><key>tile-data</key><dict><key>file-data</key><dict><key>_CFURLString</key><string>$app</string><key>_CFURLStringType</key><integer>0</integer></dict></dict></dict>"; done; killall Dock
```

## 📝 Verify Your App Paths First

App bundle names must match **exactly** or the icon won't pin. Quick check:

```bash
ls -d /Applications/zoom.us.app /Applications/Box.app "/Applications/Microsoft Outlook.app" "/Applications/Microsoft Word.app" "/Applications/Microsoft Excel.app" "/Applications/Microsoft PowerPoint.app"
```

Common gotchas on our fleet:

* **Zoom** → `zoom.us.app` (lowercase, not "Zoom.app")
* **Box Drive** → usually `Box.app` (older installs may be `Box Drive.app`)
* **Microsoft apps** → full name with space, e.g. `Microsoft Outlook.app`

## 🎨 Script Editor Version (clickable app)

Open **Script Editor** → new document → paste, then **File → Export → File Format: Application**:

```applescript
set appList to {"/Applications/zoom.us.app", "/Applications/Box.app", "/Applications/Microsoft Outlook.app", "/Applications/Microsoft Word.app", "/Applications/Microsoft Excel.app", "/Applications/Microsoft PowerPoint.app"}
do shell script "defaults write com.apple.dock persistent-apps -array"
repeat with a in appList
	do shell script "defaults write com.apple.dock persistent-apps -array-add '<dict><key>tile-data</key><dict><key>file-data</key><dict><key>_CFURLString</key><string>" & a & "</string><key>_CFURLStringType</key><integer>0</integer></dict></dict></dict>'"
end repeat
do shell script "killall Dock"
```


---


Perfect — let's fold in **Google Chrome, Self Service, Personal Print Manager, and GlobalProtect** so you get the complete standardized executive Dock in one shot. 👇

## ✅ One-Liner — Full Core App Set (native `defaults`, clear + rebuild)

```bash
defaults write com.apple.dock persistent-apps -array; for app in "/Applications/zoom.us.app" "/Applications/Box.app" "/Applications/Microsoft Outlook.app" "/Applications/Microsoft Word.app" "/Applications/Microsoft Excel.app" "/Applications/Microsoft PowerPoint.app" "/Applications/Google Chrome.app" "/Applications/Self Service.app" "/Applications/Personal Print Manager.app" "/Applications/GlobalProtect.app"; do defaults write com.apple.dock persistent-apps -array-add "<dict><key>tile-data</key><dict><key>file-data</key><dict><key>_CFURLString</key><string>$app</string><key>_CFURLStringType</key><integer>0</integer></dict></dict></dict>"; done; killall Dock
```

This empties the Dock's app side, then pins all **10 apps** in order and restarts the Dock so they appear immediately. [\[discussion....apple.com\]](https://discussions.apple.com/thread/254797028), [\[discussion....apple.com\]](https://discussions.apple.com/thread/254332755)

## 📝 Verify Paths First (these 4 are the usual troublemakers)

```bash
ls -d "/Applications/Google Chrome.app" "/Applications/Self Service.app" "/Applications/Personal Print Manager.app" "/Applications/GlobalProtect.app"
```

Common fleet gotchas:

* **Self Service (Jamf)** → normally `/Applications/Self Service.app` — but some deployments name it `Self Service+.app`
* **GlobalProtect** → usually `/Applications/GlobalProtect.app` (occasionally lives under a vendor subfolder)
* **Personal Print Manager (LRS)** → confirm the exact bundle name; sometimes it's `Personal Print Manager.app` and sometimes shortened
* **Google Chrome** → `Google Chrome.app` (with the space)

If any path comes back "No such file," fix the name before running the pin command or that one icon will silently skip.

## 🎨 Script Editor Version (clickable "Standardize Executive Dock" app)

Open **Script Editor** → new document → paste → **File → Export → File Format: Application**:

```applescript
set appList to {"/Applications/zoom.us.app", "/Applications/Box.app", "/Applications/Microsoft Outlook.app", "/Applications/Microsoft Word.app", "/Applications/Microsoft Excel.app", "/Applications/Microsoft PowerPoint.app", "/Applications/Google Chrome.app", "/Applications/Self Service.app", "/Applications/Personal Print Manager.app", "/Applications/GlobalProtect.app"}
do shell script "defaults write com.apple.dock persistent-apps -array"
repeat with a in appList
	do shell script "defaults write com.apple.dock persistent-apps -array-add '<dict><key>tile-data</key><dict><key>file-data</key><dict><key>_CFURLString</key><string>" & a & "</string><key>_CFURLStringType</key><integer>0</integer></dict></dict></dict>'"
end repeat
do shell script "killall Dock"
```

---


Got it — adding **Safari** (`/Applications/Safari.app`) and **Activity Monitor**. ⚠️ Heads-up: on Tahoe 26, Activity Monitor lives under **`/System/Applications/Utilities/`**, *not* `/Applications`, so I've used the correct path below. 👇

## ✅ One-Liner — Full 12-App Set (native `defaults`, clear + rebuild)

```bash
defaults write com.apple.dock persistent-apps -array; for app in "/Applications/Safari.app" "/Applications/Google Chrome.app" "/Applications/zoom.us.app" "/Applications/Box.app" "/Applications/Microsoft Outlook.app" "/Applications/Microsoft Word.app" "/Applications/Microsoft Excel.app" "/Applications/Microsoft PowerPoint.app" "/Applications/Self Service.app" "/Applications/Personal Print Manager.app" "/Applications/GlobalProtect.app" "/System/Applications/Utilities/Activity Monitor.app"; do defaults write com.apple.dock persistent-apps -array-add "<dict><key>tile-data</key><dict><key>file-data</key><dict><key>_CFURLString</key><string>$app</string><key>_CFURLStringType</key><integer>0</integer></dict></dict></dict>"; done; killall Dock
```

This empties the Dock's app side, then pins all **12 apps** in order and restarts the Dock so they appear immediately. [\[discussion....apple.com\]](https://discussions.apple.com/thread/254797028), [\[discussion....apple.com\]](https://discussions.apple.com/thread/254332755)

## 📝 Verify Paths First (Activity Monitor is the tricky one)

```bash
ls -d "/Applications/Safari.app" "/System/Applications/Utilities/Activity Monitor.app"
```

Key notes:

* **Safari** → `/Applications/Safari.app` ✅ (standard location)
* **Activity Monitor** → `/System/Applications/Utilities/Activity Monitor.app` — Apple moved system utilities into the read-only System volume, so `/Applications/Utilities/...` will **fail** on modern macOS. The path above is correct for Tahoe 26.

## 🎨 Script Editor Version (clickable "Standardize Executive Dock" app)

Open **Script Editor** → new document → paste → **File → Export → File Format: Application**:

```applescript
set appList to {"/Applications/Safari.app", "/Applications/Google Chrome.app", "/Applications/zoom.us.app", "/Applications/Box.app", "/Applications/Microsoft Outlook.app", "/Applications/Microsoft Word.app", "/Applications/Microsoft Excel.app", "/Applications/Microsoft PowerPoint.app", "/Applications/Self Service.app", "/Applications/Personal Print Manager.app", "/Applications/GlobalProtect.app", "/System/Applications/Utilities/Activity Monitor.app"}
do shell script "defaults write com.apple.dock persistent-apps -array"
repeat with a in appList
	do shell script "defaults write com.apple.dock persistent-apps -array-add '<dict><key>tile-data</key><dict><key>file-data</key><dict><key>_CFURLString</key><string>" & a & "</string><key>_CFURLStringType</key><integer>0</integer></dict></dict></dict>'"
end repeat
do shell script "killall Dock"
```


