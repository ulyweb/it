-- ============================================================
--  Map Network Drive
--  Version : 1.3  (auto-detects current logged-in user as default)
--  Purpose : Mount / unmount the corp SMB "files" share to
--            ~/Desktop/files with a secure GUI password prompt.
-- ============================================================

-- ---------- Configuration ----------
set smbHost to "domain.com"
set smbShare to "files/archive/scripts"
set mountPoint to (POSIX path of (path to home folder)) & "Desktop/files"

-- ---------- Detect current logged-in user (fallback to whoami) ----------
set currentUser to do shell script "stat -f%Su /dev/console 2>/dev/null || whoami"

-- ---------- Main menu ----------
set actionChoice to choose from list {"Mount Network Drive", "Unmount Network Drive"} with title "Map Network Drive" with prompt "Select an action:" default items {"Mount Network Drive"} without empty selection allowed

if actionChoice is false then return
set choice to item 1 of actionChoice

-- ============================================================
--  MOUNT
-- ============================================================
if choice is "Mount Network Drive" then
    
    -- Prompt for username (pre-filled with the current logged-in user)
    set smbUser to text returned of (display dialog "Enter your network username:" default answer currentUser with title "Map Network Drive — Username" with icon note)
    
    if smbUser is "" then
        display dialog "No username entered. Mount cancelled." buttons {"OK"} default button "OK" with icon caution
        return
    end if
    
    -- Secure password prompt (masked input)
    set pwd to text returned of (display dialog "Enter your password for " & smbUser & "@" & smbHost & ":" default answer "" with title "Map Network Drive — Authenticate" with icon note with hidden answer)
    
    if pwd is "" then
        display dialog "No password entered. Mount cancelled." buttons {"OK"} default button "OK" with icon caution
        return
    end if
    
    -- URL-encode the password using NATIVE bash only (no python/perl, no CLT)
    set encPwd to do shell script "pw=" & quoted form of pwd & "; out=''; i=0; while [ $i -lt ${#pw} ]; do c=${pw:$i:1}; case \"$c\" in [a-zA-Z0-9.~_-]) out=\"$out$c\";; *) out=\"$out$(printf '%%%02X' \"'$c\")\";; esac; i=$((i+1)); done; printf '%s' \"$out\""
    
    try
        -- Create mount point if it doesn't exist
        do shell script "mkdir -p " & quoted form of mountPoint
        
        -- Mount the share (encoded password passed inline, never shown in Terminal)
        do shell script "mount_smbfs //" & smbUser & ":" & encPwd & "@" & smbHost & "/" & smbShare & " " & quoted form of mountPoint
        
        display dialog "✅ Network drive mounted successfully." & return & return & "User: " & smbUser & return & "Location: " & mountPoint buttons {"Open in Finder", "OK"} default button "OK" with title "Map Network Drive — Complete"
        if button returned of result is "Open in Finder" then
            do shell script "open " & quoted form of mountPoint
        end if
        
    on error errMsg
        display dialog "❌ Mount failed:" & return & return & errMsg buttons {"OK"} default button "OK" with icon stop with title "Map Network Drive — Error"
    end try
    
-- ============================================================
--  UNMOUNT
-- ============================================================
else if choice is "Unmount Network Drive" then
    
    try
        do shell script "umount " & quoted form of mountPoint
        display dialog "✅ Network drive unmounted successfully." & return & return & "Location: " & mountPoint buttons {"OK"} default button "OK" with title "Map Network Drive — Complete"
    on error errMsg
        display dialog "❌ Unmount failed (the drive may not be mounted):" & return & return & errMsg buttons {"OK"} default button "OK" with icon stop with title "Map Network Drive — Error"
    end try
    
end if
