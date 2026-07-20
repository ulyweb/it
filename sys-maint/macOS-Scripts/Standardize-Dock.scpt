-- ============================================================
--  Standardize Dock — Interactive Chooser
--  Option 1: Standard Core Apps Dock (full 12-app set)
--  Option 2: Fully reset Dock to macOS factory defaults
-- ============================================================

set appList to {"/Applications/Safari.app", "/Applications/Google Chrome.app", "/Applications/zoom.us.app", "/Applications/Box.app", "/Applications/Microsoft Outlook.app", "/Applications/Microsoft Word.app", "/Applications/Microsoft Excel.app", "/Applications/Microsoft PowerPoint.app", "/Applications/Self Service.app", "/Applications/Personal Print Manager.app", "/Applications/GlobalProtect.app", "/System/Applications/Utilities/Activity Monitor.app"}

set actionChoice to choose from list {"Standard Core Apps Dock", "Fully Reset Dock to Factory Defaults"} with title "Standardize Dock" with prompt "Select the Dock action to perform:" default items {"Standard Core Apps Dock"} without empty selection allowed

if actionChoice is false then
    -- User clicked Cancel
    return
end if

set choice to item 1 of actionChoice

if choice is "Standard Core Apps Dock" then
    -- Clear the app side, then pin the approved 12-app set
    do shell script "defaults write com.apple.dock persistent-apps -array"
    repeat with a in appList
        do shell script "defaults write com.apple.dock persistent-apps -array-add '<dict><key>tile-data</key><dict><key>file-data</key><dict><key>_CFURLString</key><string>" & a & "</string><key>_CFURLStringType</key><integer>0</integer></dict></dict></dict>'"
    end repeat
    do shell script "killall Dock"
    display notification "Standard core apps Dock applied." with title "Standardize Dock" subtitle "Complete"
    
else if choice is "Fully Reset Dock to Factory Defaults" then
    -- Wipe all custom Dock settings; macOS regenerates factory defaults
    do shell script "defaults delete com.apple.dock; killall Dock"
    display notification "Dock reset to macOS factory defaults." with title "Standardize Dock" subtitle "Complete"
end if
