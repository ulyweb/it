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


