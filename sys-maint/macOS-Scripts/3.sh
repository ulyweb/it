#!/bin/bash
#===============================================================================
# MAC-HANDOFF v3.0  —  Cryptographic Hand-Off Gatekeeper + LIVE TELEMETRY CONSOLE
#
# *** LAUNCH WITHOUT SUDO ***      ./crypto-handoff_v3.0.sh
#
# v3.0 UPGRADE — worldmonitor.app / PG&E Power Monitor style console:
#   • Fixed command bar: brand + version + LIVE pulse + GO/NO-GO state badge
#   • ⌘K command palette (search gates, evidence, actions)
#   • Icon toolbar: ◐ theme · ⟳ refresh · ⤓ EXPORT (JSON/CSV) · ⚙ SETTINGS · ⛶ fullscreen
#   • Live ticking clock (local + UTC) with session uptime
#   • VIEW chips: OVERVIEW / GATES / EVIDENCE / JAMF / COPILOT
#   • Dismissible priority alert banner on NO-GO
#   • Hero verdict card (big GO / NO-GO) + countdown-style posture line
#   • KPI tile strip with hover-lift + gold glow
#   • Collapsible panels with ? info tooltips and ✕ close, UNAVAILABLE states
#   • Copilot diagnostic prompts with copy-to-clipboard
#   • All data injected as a single JSON payload -> JS renders the console
#
# Automates Phase 3 (Audit) + Phase 4 (Cleanup) of the Mac Provisioning guide.
# Runs LOCALLY. Jamf is used view-only afterward to confirm the recovery key +
# owner by username + serial.
#
# Refuses to delete the Staging Admin unless the Primary User holds Secure Token,
# Volume Ownership, FileVault is On, is FileVault-enabled, Bootstrap Token is
# escrowed, and a surviving token holder remains (anti-"orphaned disk" gate).
#
# Family: MAC-MAINT toolkit • bash 3.2-safe • zero runtime dependencies
#
# USAGE (run WITHOUT sudo):
#   ./crypto-handoff_v3.0.sh                 # fully interactive (GUI password)
#   ./crypto-handoff_v3.0.sh --audit-only    # Phase 3 only; CANNOT delete
#   ./crypto-handoff_v3.0.sh --dry-run       # walks Phase 4 logic, prints only
#   ./crypto-handoff_v3.0.sh --reissue-prk   # rotate + re-escrow a fresh PRK
#   ./crypto-handoff_v3.0.sh --no-reboot     # skip the Phase 5 cold-boot reboot
#   ./crypto-handoff_v3.0.sh --help
#===============================================================================

set -u
SCRIPT_VERSION="3.0"
TIMESTAMP="$(date '+%Y%m%d_%H%M%S')"
NOW_HUMAN="$(date '+%b %d, %Y %I:%M %p %Z')"
NOW_ISO="$(date -u '+%Y-%m-%dT%H:%M:%SZ')"
SERIAL="$(system_profiler SPHardwareDataType 2>/dev/null | awk '/Serial Number/{print $NF; exit}')"
[ -z "$SERIAL" ] && SERIAL="UNKNOWN"
HOST="$(scutil --get ComputerName 2>/dev/null || hostname)"
MODEL="$(system_profiler SPHardwareDataType 2>/dev/null | awk -F': ' '/Model Name/{print $2; exit}')"
[ -z "$MODEL" ] && MODEL="Mac"
CHIP="$(system_profiler SPHardwareDataType 2>/dev/null | awk -F': ' '/Chip|Processor Name/{print $2; exit}')"
[ -z "$CHIP" ] && CHIP="unknown"
OSVER="$(sw_vers -productVersion 2>/dev/null)"
OSBUILD="$(sw_vers -buildVersion 2>/dev/null)"
OWNER_PW=""
AUTH_METHOD="none"
KEEPALIVE_PID=""
START_EPOCH="$(date +%s)"

# ---- Colors -----------------------------------------------------------------
if [ -t 1 ]; then
  C_G=$'\033[0;32m'; C_R=$'\033[0;31m'; C_Y=$'\033[0;33m'; C_B=$'\033[0;36m'; C_D=$'\033[1m'; C_0=$'\033[0m'
else
  C_G=""; C_R=""; C_Y=""; C_B=""; C_D=""; C_0=""
fi

# ---- Flags ------------------------------------------------------------------
AUDIT_ONLY=0; DRY_RUN=0; REISSUE_PRK=0; DO_REBOOT=1
for arg in "$@"; do
  case "$arg" in
    --audit-only) AUDIT_ONLY=1 ;;
    --dry-run)    DRY_RUN=1 ;;
    --reissue-prk) REISSUE_PRK=1 ;;
    --no-reboot)  DO_REBOOT=0 ;;
    --help|-h)    grep '^#' "$0" | sed 's/^# \{0,1\}//'; exit 0 ;;
    *) echo "${C_Y}Ignoring unknown argument: $arg${C_0}" ;;
  esac
done

# ---- Identity / session -----------------------------------------------------
CONSOLE_USER="$(/usr/bin/stat -f%Su /dev/console 2>/dev/null)"
CONSOLE_UID="$(id -u "$CONSOLE_USER" 2>/dev/null)"
STARTED_AS_ROOT=0; [ "$(id -u)" -eq 0 ] && STARTED_AS_ROOT=1
LOGIN_USER="$(logname 2>/dev/null)"; [ -z "$LOGIN_USER" ] && LOGIN_USER="${SUDO_USER:-$CONSOLE_USER}"

# ---- Output dir (self-healing) ----------------------------------------------
BASE_DIR="/Users/Shared/MAC-HANDOFF"
if [ -d "$BASE_DIR" ] && [ ! -w "$BASE_DIR" ]; then
  if [ "$(id -u)" -eq 0 ]; then
    chown -R "${LOGIN_USER:-root}" "$BASE_DIR" 2>/dev/null
  fi
fi
OUTDIR="${BASE_DIR}/${TIMESTAMP}_${SERIAL}"
mkdir -p "$OUTDIR" 2>/dev/null
if [ ! -d "$OUTDIR" ] || [ ! -w "$OUTDIR" ]; then
  OUTDIR="$HOME/MAC-HANDOFF/${TIMESTAMP}_${SERIAL}"; mkdir -p "$OUTDIR" 2>/dev/null
fi
if [ ! -d "$OUTDIR" ] || [ ! -w "$OUTDIR" ]; then
  OUTDIR="$(mktemp -d "/tmp/MAC-HANDOFF_${SERIAL}.XXXXXX")"
fi
LOG="${OUTDIR}/handoff.log"
HTML="${OUTDIR}/MAC-HANDOFF_Console_${TIMESTAMP}.html"
JSONF="${OUTDIR}/handoff_payload.json"
if ! : > "$LOG" 2>/dev/null; then
  echo "${C_R}FATAL: Cannot create log file in $OUTDIR — check permissions.${C_0}" >&2
  exit 6
fi
chmod -R 0777 "$OUTDIR" 2>/dev/null || true

# ---- Logging ----------------------------------------------------------------
log(){ echo "$(date '+%H:%M:%S')  $1" >> "$LOG" 2>/dev/null; }
banner(){ echo ""; echo "${C_B}${C_D}==== $1 ====${C_0}"; log "==== $1 ===="; }
ok(){    echo "  ${C_G}[ PASS ]${C_0} $1"; log "[PASS] $1"; }
fail(){  echo "  ${C_R}[ FAIL ]${C_0} $1"; log "[FAIL] $1"; }
warn(){  echo "  ${C_Y}[ WARN ]${C_0} $1"; log "[WARN] $1"; }
info(){  echo "  ${C_B}[ .... ]${C_0} $1"; log "[INFO] $1"; }

cleanup(){
  [ -n "${KEEPALIVE_PID:-}" ] && kill "$KEEPALIVE_PID" >/dev/null 2>&1
  OWNER_PW=""; unset OWNER_PW
}
trap cleanup EXIT

# ---- JSON helper: escape a string --------------------------------------------
jesc(){ printf '%s' "$1" | sed -e 's/\\/\\\\/g' -e 's/"/\\"/g' | tr '\n' ' ' | sed 's/\t/ /g'; }

# ---- GUI helpers -------------------------------------------------------------
run_osa(){
  if [ "$(id -u)" -eq 0 ] && [ -n "$CONSOLE_UID" ] && [ "$CONSOLE_USER" != "root" ]; then
    launchctl asuser "$CONSOLE_UID" sudo -u "$CONSOLE_USER" /usr/bin/osascript "$@"
  else
    /usr/bin/osascript "$@"
  fi
}
ui_password(){
  run_osa \
    -e "try" \
    -e "set r to display dialog \"$2\" with title \"$1\" default answer \"\" with icon note buttons {\"Cancel\",\"OK\"} default button \"OK\" cancel button \"Cancel\" with hidden answer" \
    -e "return text returned of r" \
    -e "on error number -128" \
    -e "return \"___CANCELLED___\"" \
    -e "end try" 2>/dev/null
}
ui_alert(){ run_osa -e "display dialog \"$2\" with title \"$1\" buttons {\"OK\"} default button \"OK\" with icon caution" >/dev/null 2>&1; }

priv(){ if [ "$(id -u)" -eq 0 ]; then "$@"; else sudo -n "$@"; fi; }

# ---- Header -----------------------------------------------------------------
MODE_LABEL="AUDIT + CLEANUP"
[ $DRY_RUN -eq 1 ]    && MODE_LABEL="DRY-RUN"
[ $AUDIT_ONLY -eq 1 ] && MODE_LABEL="AUDIT ONLY"

clear 2>/dev/null
echo ""
echo "${C_D}  ================================================================${C_0}"
echo "${C_D}   MAC-HANDOFF v${SCRIPT_VERSION}  —  Cryptographic Hand-Off Console${C_0}"
echo "${C_D}  ================================================================${C_0}"
echo "   Host:    ${HOST}  (${MODEL} · ${CHIP})"
echo "   Serial:  ${SERIAL}"
echo "   macOS:   ${OSVER} (${OSBUILD})"
echo "   Time:    ${NOW_HUMAN}"
echo "   Mode:    ${C_Y}${MODE_LABEL}${C_0}"
echo "   Output:  ${OUTDIR}"
echo ""

# ---- Account prompts ---------------------------------------------------------
prompt_primary(){
  local input=""
  echo "${C_B}Who is the PRIMARY USER (the device owner who must keep access)?${C_0}"
  if [ -n "$CONSOLE_USER" ] && [ "$CONSOLE_USER" != "root" ]; then
    printf "  Enter username [Enter = logged-in: %s]: " "$CONSOLE_USER"
    read -r input; [ -z "$input" ] && input="$CONSOLE_USER"
  else
    printf "  Enter Primary User's short username: "; read -r input
  fi
  PRIMARY_USER="$input"
}
prompt_staging(){
  local input=""
  echo ""
  echo "${C_B}Which STAGING ADMIN account should be removed after the hand-off?${C_0}"
  echo "  (Local admin accounts currently on this Mac:)"
  dscl . -list /Users | grep -v '^_' | grep -vE '^(daemon|nobody|root)$' | while read -r u; do
    if dseditgroup -o checkmember -m "$u" admin >/dev/null 2>&1; then echo "    ${C_Y}- $u  (admin)${C_0}"; fi
  done
  printf "  Enter Staging Admin's short username: "; read -r input
  STAGING_ADMIN="$input"
}

banner "STEP 0 - IDENTIFY ACCOUNTS"
PRIMARY_USER=""; STAGING_ADMIN=""
prompt_primary
[ $AUDIT_ONLY -eq 0 ] && prompt_staging

echo ""
echo "  ${C_D}Confirm:${C_0}"
echo "    Primary User  : ${C_G}${PRIMARY_USER:-<none>}${C_0}"
[ $AUDIT_ONLY -eq 0 ] && echo "    Staging Admin : ${C_R}${STAGING_ADMIN:-<none>}${C_0}  (to be removed)"
printf "  Proceed with these accounts? [y/N]: "
read -r CONFIRM
case "$CONFIRM" in
  y|Y|yes|YES) ok "Confirmed. Continuing." ;;
  *) echo "${C_Y}Cancelled by operator. No changes made.${C_0}"; exit 0 ;;
esac

# ---- Preflight ---------------------------------------------------------------
banner "PREFLIGHT"
PREFLIGHT_OK=1
if [ -z "$PRIMARY_USER" ]; then fail "Primary User not provided."; PREFLIGHT_OK=0
elif ! id "$PRIMARY_USER" >/dev/null 2>&1; then fail "Primary User '$PRIMARY_USER' does not exist."; PREFLIGHT_OK=0
else ok "Primary User account exists: $PRIMARY_USER"; fi
if [ $AUDIT_ONLY -eq 0 ]; then
  if [ -z "$STAGING_ADMIN" ]; then fail "Staging Admin not provided."; PREFLIGHT_OK=0
  elif ! id "$STAGING_ADMIN" >/dev/null 2>&1; then fail "Staging Admin '$STAGING_ADMIN' does not exist."; PREFLIGHT_OK=0
  elif [ "$STAGING_ADMIN" = "$PRIMARY_USER" ]; then fail "Staging Admin equals Primary User - refusing."; PREFLIGHT_OK=0
  else ok "Staging Admin account exists: $STAGING_ADMIN"; fi
fi
[ $PREFLIGHT_OK -eq 0 ] && { echo ""; echo "${C_R}Preflight failed. Aborting.${C_0}"; exit 3; }

#===============================================================================
# AUTHENTICATION
#===============================================================================
banner "AUTHENTICATION"
RUN_USER="$CONSOLE_USER"; [ -z "$RUN_USER" ] && RUN_USER="$LOGIN_USER"

if [ $STARTED_AS_ROOT -eq 1 ]; then
  warn "You launched WITH sudo. Next time run WITHOUT sudo for a fully GUI flow."
  ok "Root privileges already present."
  AUTH_METHOD="root(sudo-launch)"
else
  if ! dseditgroup -o checkmember -m "$RUN_USER" admin >/dev/null 2>&1; then
    fail "$RUN_USER is NOT a local admin - cannot elevate."
    ui_alert "MAC-HANDOFF" "The logged-in user ($RUN_USER) is not an administrator. Run 'Make Me Admin' from Self Service, then re-run this tool."
    exit 5
  fi
  info "Requesting administrator password via secure GUI dialog (masked dots)..."
  attempts=0; primed=0
  while [ $attempts -lt 3 ]; do
    OWNER_PW="$(ui_password "MAC-HANDOFF - Administrator Password" "Enter the login password for $RUN_USER to authorize this secure device hand-off. Dots appear as you type so you can confirm your entry.")"
    if [ "$OWNER_PW" = "___CANCELLED___" ]; then OWNER_PW=""; echo "${C_Y}Cancelled at password dialog. No changes made.${C_0}"; exit 0; fi
    if printf '%s\n' "$OWNER_PW" | sudo -S -p '' -v 2>/dev/null; then primed=1; break; fi
    attempts=$((attempts+1)); OWNER_PW=""
    [ $attempts -lt 3 ] && ui_alert "MAC-HANDOFF" "That password was incorrect. Attempts remaining: $((3-attempts))"
  done
  if [ $primed -ne 1 ]; then fail "Password incorrect after 3 attempts. Aborting."; exit 4; fi
  ok "Password validated via GUI. Sudo primed - no further Terminal prompts."
  ( while true; do sudo -n -v >/dev/null 2>&1; sleep 50; done ) & KEEPALIVE_PID=$!
  AUTH_METHOD="gui"
fi

ensure_owner_credential(){
  [ -n "$OWNER_PW" ] && return 0
  local acct="$RUN_USER" a=0
  while [ $a -lt 3 ]; do
    OWNER_PW="$(ui_password "MAC-HANDOFF - Authorize Hand-Off" "Enter the login password for $acct to authorize removing the staging admin. Dots appear as you type.")"
    [ "$OWNER_PW" = "___CANCELLED___" ] && { OWNER_PW=""; return 2; }
    if dscl /Local/Default -authonly "$acct" "$OWNER_PW" >/dev/null 2>&1; then return 0; fi
    a=$((a+1)); OWNER_PW=""
    [ $a -lt 3 ] && ui_alert "MAC-HANDOFF" "That password was incorrect. Attempts remaining: $((3-a))"
  done
  return 1
}

guid_of(){ dscl . -read "/Users/$1" GeneratedUID 2>/dev/null | awk '{print $2}'; }
PRIMARY_GUID="$(guid_of "$PRIMARY_USER")"

# ---- Gate tracking (parallel arrays, bash 3.2-safe) --------------------------
G_ID=(); G_NAME=(); G_STATUS=(); G_DETAIL=(); G_CMD=(); G_WHY=()
add_gate(){ G_ID+=("$1"); G_NAME+=("$2"); G_STATUS+=("$3"); G_DETAIL+=("$4"); G_CMD+=("$5"); G_WHY+=("$6"); }

#===============================================================================
# PHASE 3 - CRYPTOGRAPHIC AUDIT
#===============================================================================
banner "PHASE 3 - CRYPTOGRAPHIC AUDIT"

info "Checking FileVault status..."
FV_RAW="$(priv fdesetup status 2>&1)"
if echo "$FV_RAW" | grep -q "FileVault is On"; then
  ok "FileVault is On"
  add_gate "fv" "FileVault Encryption" "PASS" "$FV_RAW" "fdesetup status" "The volume must be encrypted before ownership can be transferred safely."
else
  fail "FileVault is NOT on"
  add_gate "fv" "FileVault Encryption" "FAIL" "$FV_RAW" "fdesetup status" "The volume must be encrypted before ownership can be transferred safely."
fi

info "Checking Secure Token for $PRIMARY_USER..."
ST_RAW="$(priv sysadminctl -secureTokenStatus "$PRIMARY_USER" 2>&1)"
CRYPTO_HIT=""
[ -n "$PRIMARY_GUID" ] && CRYPTO_HIT="$(diskutil apfs listCryptoUsers / 2>/dev/null | grep -i "$PRIMARY_GUID")"
if echo "$ST_RAW" | grep -q "ENABLED"; then
  if [ -n "$CRYPTO_HIT" ]; then
    ok "Secure Token ENABLED (confirmed in APFS crypto users)"
    add_gate "token" "Secure Token (Primary)" "PASS" "sysadminctl=ENABLED; APFS crypto-user GUID matched" "sysadminctl -secureTokenStatus + diskutil apfs listCryptoUsers /" "Without a Secure Token the owner cannot unlock FileVault after the staging admin is removed."
  else
    warn "sysadminctl=ENABLED but GUID not matched in crypto users"
    add_gate "token" "Secure Token (Primary)" "WARN" "ENABLED but no diskutil cross-check" "sysadminctl -secureTokenStatus" "Without a Secure Token the owner cannot unlock FileVault after the staging admin is removed."
  fi
else
  fail "Secure Token DISABLED / unknown for $PRIMARY_USER"
  add_gate "token" "Secure Token (Primary)" "FAIL" "$ST_RAW" "sysadminctl -secureTokenStatus" "Without a Secure Token the owner cannot unlock FileVault after the staging admin is removed."
fi

info "Checking Volume Ownership for $PRIMARY_USER..."
VO="$(diskutil apfs listUsers / 2>/dev/null | awk -v id="$PRIMARY_GUID" '
  toupper($0) ~ toupper(id){found=1} found && /Volume Owner:/{print $NF; exit}')"
if [ "$VO" = "Yes" ]; then
  ok "Volume Owner: Yes"
  add_gate "owner" "Volume Ownership" "PASS" "Volume Owner: Yes" "diskutil apfs listUsers /" "Volume owners can authorize system updates and recovery operations."
else
  fail "Volume Owner: ${VO:-No/Unknown}"
  add_gate "owner" "Volume Ownership" "FAIL" "Volume Owner: ${VO:-Unknown}" "diskutil apfs listUsers /" "Volume owners can authorize system updates and recovery operations."
fi

info "Checking Bootstrap Token escrow..."
BT_RAW="$(priv profiles status -type bootstraptoken 2>&1)"
if echo "$BT_RAW" | grep -qi "escrowed to server: YES"; then
  ok "Bootstrap Token escrowed to server: YES"
  add_gate "bootstrap" "Bootstrap Token Escrow" "PASS" "$BT_RAW" "profiles status -type bootstraptoken" "Lets MDM grant Secure Tokens later and authorize updates without the staging admin."
else
  fail "Bootstrap Token NOT escrowed to server"
  add_gate "bootstrap" "Bootstrap Token Escrow" "FAIL" "$BT_RAW" "profiles status -type bootstraptoken" "Lets MDM grant Secure Tokens later and authorize updates without the staging admin."
fi

info "Checking FileVault-enabled users..."
FVL_RAW="$(priv fdesetup list 2>&1)"
if echo "$FVL_RAW" | grep -qi "^${PRIMARY_USER},"; then
  ok "$PRIMARY_USER is FileVault-enabled"
  add_gate "fvuser" "FileVault-Enabled User" "PASS" "$FVL_RAW" "fdesetup list" "The owner must appear at the pre-boot login screen or the Mac becomes unusable."
else
  fail "$PRIMARY_USER not found in fdesetup list"
  add_gate "fvuser" "FileVault-Enabled User" "FAIL" "$FVL_RAW" "fdesetup list" "The owner must appear at the pre-boot login screen or the Mac becomes unusable."
fi

GATES_FAILED=0; GATES_WARN=0; GATES_PASS=0
i=0; while [ $i -lt ${#G_STATUS[@]} ]; do
  case "${G_STATUS[$i]}" in
    FAIL) GATES_FAILED=$((GATES_FAILED+1)) ;;
    WARN) GATES_WARN=$((GATES_WARN+1)) ;;
    PASS) GATES_PASS=$((GATES_PASS+1)) ;;
  esac
  i=$((i+1))
done
GATES_TOTAL=${#G_STATUS[@]}

echo ""
if [ $GATES_FAILED -gt 0 ]; then
  echo "${C_R}${C_D}  DECISION: NO-GO  - ${GATES_FAILED} gate(s) failed. Staging Admin will NOT be removed.${C_0}"
  DECISION="NO-GO"
else
  echo "${C_G}${C_D}  DECISION: GO  - all cryptographic gates passed.${C_0}"
  DECISION="GO"
fi

#===============================================================================
# Optional PRK reissue
#===============================================================================
PRK_RESULT="Not requested"
if [ "$DECISION" = "GO" ] && [ $REISSUE_PRK -eq 1 ] && [ $AUDIT_ONLY -eq 0 ]; then
  banner "PRK REISSUE"
  if [ $DRY_RUN -eq 1 ]; then
    warn "DRY-RUN: would run fdesetup changerecovery -personal"
    PRK_RESULT="Dry-run (not executed)"
  else
    ensure_owner_credential; EC=$?
    if [ $EC -eq 0 ] && [ -n "$OWNER_PW" ]; then
      if priv fdesetup changerecovery -personal -inputplist >>"$LOG" 2>&1 <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
<key>Username</key><string>${RUN_USER}</string>
<key>Password</key><string>${OWNER_PW}</string>
</dict></plist>
PLIST
      then ok "New Personal Recovery Key generated"; PRK_RESULT="Rotated + escrow pending recon"
      else fail "PRK reissue failed (see log)"; PRK_RESULT="Failed (see log)"; fi
    else warn "No owner credential - skipping PRK reissue."; PRK_RESULT="Skipped (no credential)"; fi
  fi
fi

#===============================================================================
# PHASE 4 - CLEANUP & RECON
#===============================================================================
DELETE_RESULT="Skipped"
SURVIVOR="none"
PREBOOT_RESULT="not run"
RECON_RESULT="not run"

if [ $AUDIT_ONLY -eq 1 ]; then
  banner "PHASE 4 - SKIPPED (audit-only mode)"; DELETE_RESULT="Skipped (audit-only)"
elif [ "$DECISION" != "GO" ]; then
  banner "PHASE 4 - SKIPPED (NO-GO decision)"; DELETE_RESULT="Skipped (NO-GO)"
else
  banner "PHASE 4 - CLEANUP & RECON"
  info "Verifying a surviving Secure Token holder will remain..."
  OTHER_TOKEN=0
  for u in $(dscl . -list /Users | grep -v '^_' | grep -vE '^(daemon|nobody|root)$'); do
    [ "$u" = "$STAGING_ADMIN" ] && continue
    if priv sysadminctl -secureTokenStatus "$u" 2>&1 | grep -q "ENABLED"; then OTHER_TOKEN=1; SURVIVOR="$u"; break; fi
  done

  if [ $OTHER_TOKEN -eq 0 ]; then
    fail "No surviving Secure Token holder - refusing to delete (would orphan disk)."
    DELETE_RESULT="Refused (no surviving token holder)"
    add_gate "survivor" "Surviving Token Holder" "FAIL" "$DELETE_RESULT" "sysadminctl -secureTokenStatus (all users)" "At least one Secure Token holder must remain or the disk is orphaned."
    GATES_FAILED=$((GATES_FAILED+1)); GATES_TOTAL=$((GATES_TOTAL+1))
  else
    ok "Surviving Secure Token holder confirmed: $SURVIVOR"
    add_gate "survivor" "Surviving Token Holder" "PASS" "Survivor: $SURVIVOR" "sysadminctl -secureTokenStatus (all users)" "At least one Secure Token holder must remain or the disk is orphaned."
    GATES_PASS=$((GATES_PASS+1)); GATES_TOTAL=$((GATES_TOTAL+1))

    if [ $DRY_RUN -eq 1 ]; then
      echo ""
      warn "DRY-RUN - would execute (NOT run now):"
      echo "      ${C_D}sysadminctl -deleteUser \"$STAGING_ADMIN\" -secure -adminUser \"$RUN_USER\" -adminPassword <owner-pw>${C_0}"
      DELETE_RESULT="Dry-run (command printed)"
      PREBOOT_RESULT="dry-run"; RECON_RESULT="dry-run"
    else
      info "Deleting Staging Admin: $STAGING_ADMIN ..."
      if [ -z "$OWNER_PW" ]; then ensure_owner_credential; fi
      if [ -n "$OWNER_PW" ]; then
        priv sysadminctl -deleteUser "$STAGING_ADMIN" -secure -adminUser "$RUN_USER" -adminPassword "$OWNER_PW" 2>>"$LOG"
      else
        warn "No GUI credential - falling back to Terminal interactive auth."
        priv sysadminctl -deleteUser "$STAGING_ADMIN" -secure interactive 2>>"$LOG"
        AUTH_METHOD="terminal"
      fi
      if ! id "$STAGING_ADMIN" >/dev/null 2>&1; then ok "Staging Admin removed"; DELETE_RESULT="Deleted"
      else fail "Delete ran but account still present (see log)"; DELETE_RESULT="Delete incomplete"; fi

      info "Updating APFS preboot volume..."
      if priv diskutil apfs updatePreboot / >>"$LOG" 2>&1; then ok "Preboot updated"; PREBOOT_RESULT="updated"
      else warn "updatePreboot returned non-zero"; PREBOOT_RESULT="non-zero exit"; fi

      info "Forcing Jamf inventory sync (jamf recon)..."
      if command -v jamf >/dev/null 2>&1; then
        if priv jamf recon >>"$LOG" 2>&1; then ok "jamf recon complete"; RECON_RESULT="complete"
        else warn "jamf recon returned non-zero"; RECON_RESULT="non-zero exit"; fi
      else warn "jamf binary not found"; RECON_RESULT="jamf binary not found"; fi
    fi
  fi
fi

OWNER_PW=""; unset OWNER_PW
END_EPOCH="$(date +%s)"
DURATION=$(( END_EPOCH - START_EPOCH ))

#===============================================================================
# BUILD JSON PAYLOAD
#===============================================================================
GATES_JSON=""
i=0; while [ $i -lt ${#G_ID[@]} ]; do
  [ $i -gt 0 ] && GATES_JSON="${GATES_JSON},"
  GATES_JSON="${GATES_JSON}{\"id\":\"$(jesc "${G_ID[$i]}")\",\"name\":\"$(jesc "${G_NAME[$i]}")\",\"status\":\"$(jesc "${G_STATUS[$i]}")\",\"detail\":\"$(jesc "${G_DETAIL[$i]}")\",\"cmd\":\"$(jesc "${G_CMD[$i]}")\",\"why\":\"$(jesc "${G_WHY[$i]}")\"}"
  i=$((i+1))
done

case "$AUTH_METHOD" in
  gui) AUTH_LABEL="GUI dialog (masked, validated)";;
  terminal) AUTH_LABEL="Terminal interactive (fallback)";;
  *) AUTH_LABEL="Root (launched with sudo)";;
esac

cat > "$JSONF" <<JSONEOF
{
  "version": "${SCRIPT_VERSION}",
  "generated": "${NOW_HUMAN}",
  "generatedISO": "${NOW_ISO}",
  "timestamp": "${TIMESTAMP}",
  "mode": "${MODE_LABEL}",
  "decision": "${DECISION}",
  "host": "$(jesc "$HOST")",
  "model": "$(jesc "$MODEL")",
  "chip": "$(jesc "$CHIP")",
  "serial": "${SERIAL}",
  "osver": "${OSVER}",
  "osbuild": "${OSBUILD}",
  "primaryUser": "$(jesc "$PRIMARY_USER")",
  "stagingAdmin": "$(jesc "${STAGING_ADMIN:-N/A}")",
  "survivor": "$(jesc "$SURVIVOR")",
  "authMethod": "$(jesc "$AUTH_LABEL")",
  "gatesTotal": ${GATES_TOTAL},
  "gatesPass": ${GATES_PASS},
  "gatesWarn": ${GATES_WARN},
  "gatesFail": ${GATES_FAILED},
  "deleteResult": "$(jesc "$DELETE_RESULT")",
  "prebootResult": "$(jesc "$PREBOOT_RESULT")",
  "reconResult": "$(jesc "$RECON_RESULT")",
  "prkResult": "$(jesc "$PRK_RESULT")",
  "durationSec": ${DURATION},
  "outdir": "$(jesc "$OUTDIR")",
  "gates": [${GATES_JSON}]
}
JSONEOF

#===============================================================================
# BUILD THE LIVE TELEMETRY CONSOLE (HTML)
#===============================================================================
PAYLOAD="$(cat "$JSONF")"

cat > "$HTML" <<'HTMLHEAD'
<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>MAC-HANDOFF — Cryptographic Hand-Off Console</title>
<style>
:root{--gold:#d4af37;--gold2:#f0d774;--bg:#070910;--panel:rgba(255,255,255,.045);
 --line:rgba(255,255,255,.09);--ok:#39d98a;--warn:#f0b429;--bad:#ff5c6c;--info:#6ea8ff;--txt:#e7ebf3;--dim:#9aa4b8;}
*{box-sizing:border-box;margin:0;padding:0;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;}
html{scroll-behavior:smooth;}
body{background:var(--bg);color:var(--txt);overflow-x:hidden;padding-top:104px;}
body.light{--bg:#f4f6fb;--panel:rgba(0,0,0,.03);--line:rgba(0,0,0,.10);--txt:#12161f;--dim:#5a6478;}
.bgfx{position:fixed;inset:0;z-index:0;pointer-events:none;background:
 radial-gradient(1200px 600px at 15% -10%,rgba(212,175,55,.10),transparent),
 radial-gradient(900px 500px at 110% 10%,rgba(80,140,255,.07),transparent),
 radial-gradient(700px 500px at 50% 120%,rgba(212,175,55,.05),transparent);}
.particles{position:fixed;inset:0;z-index:0;overflow:hidden;pointer-events:none;}
.particles span{position:absolute;bottom:-10px;width:5px;height:5px;background:var(--gold);border-radius:50%;
 opacity:.3;filter:drop-shadow(0 0 6px var(--gold));animation:float linear infinite;}
@keyframes float{to{transform:translateY(-110vh) translateX(40px);opacity:0;}}

/* ===== COMMAND BAR ===== */
.cmdbar{position:fixed;top:0;left:0;right:0;z-index:50;background:rgba(7,9,16,.86);
 backdrop-filter:blur(18px);border-bottom:1px solid var(--line);}
body.light .cmdbar{background:rgba(244,246,251,.9);}
.cmdrow{display:flex;align-items:center;gap:14px;padding:10px 18px;flex-wrap:wrap;}
.brand{display:flex;align-items:center;gap:8px;font-weight:900;font-size:15px;letter-spacing:.5px;white-space:nowrap;}
.brand .b1{color:var(--txt);} .brand .b2{color:var(--gold);}
.vtag{font-size:10px;font-weight:800;color:var(--gold);border:1px solid rgba(212,175,55,.45);
 padding:2px 7px;border-radius:6px;background:rgba(212,175,55,.10);}
.live{display:flex;align-items:center;gap:5px;font-size:10px;font-weight:800;letter-spacing:1px;
 color:var(--ok);border:1px solid rgba(57,217,138,.4);background:rgba(57,217,138,.12);padding:3px 9px;border-radius:999px;}
.live .dot{width:6px;height:6px;border-radius:50%;background:var(--ok);animation:pulse 1.6s ease-in-out infinite;}
@keyframes pulse{0%,100%{opacity:1;transform:scale(1);}50%{opacity:.35;transform:scale(.75);}}
.statebadge{font-size:10px;font-weight:900;letter-spacing:1.5px;padding:4px 11px;border-radius:999px;white-space:nowrap;}
.statebadge.go{color:var(--ok);border:1px solid rgba(57,217,138,.5);background:rgba(57,217,138,.14);}
.statebadge.nogo{color:var(--bad);border:1px solid rgba(255,92,108,.5);background:rgba(255,92,108,.14);}
.modebadge{font-size:10px;font-weight:800;letter-spacing:1px;color:var(--info);
 border:1px solid rgba(110,168,255,.4);background:rgba(110,168,255,.12);padding:3px 9px;border-radius:999px;}
.spacer{flex:1;}
.searchbox{display:flex;align-items:center;gap:7px;border:1px solid var(--line);background:var(--panel);
 border-radius:9px;padding:6px 11px;min-width:210px;cursor:pointer;transition:.2s;}
.searchbox:hover{border-color:rgba(212,175,55,.45);}
.searchbox span{font-size:12px;color:var(--dim);}
.kbd{font-size:10px;border:1px solid var(--line);border-radius:4px;padding:1px 5px;color:var(--dim);font-family:ui-monospace,Menlo,monospace;}
.iconbtn{border:1px solid var(--line);background:var(--panel);color:var(--txt);border-radius:8px;
 padding:6px 10px;font-size:12px;cursor:pointer;transition:.2s;white-space:nowrap;}
.iconbtn:hover{border-color:var(--gold);color:var(--gold);transform:translateY(-1px);}
.clock{font-family:ui-monospace,Menlo,monospace;font-size:12px;color:var(--gold);white-space:nowrap;}
.clock small{color:var(--dim);font-family:-apple-system,sans-serif;}

/* second row: view chips */
.cmdrow2{display:flex;align-items:center;gap:8px;padding:0 18px 10px;flex-wrap:wrap;}
.lbl{font-size:10px;color:var(--dim);letter-spacing:1.5px;font-weight:700;}
.chip{font-size:11px;font-weight:700;padding:5px 13px;border-radius:8px;cursor:pointer;
 border:1px solid var(--line);background:var(--panel);color:var(--dim);transition:.2s;}
.chip:hover{color:var(--txt);transform:translateY(-1px);}
.chip.on{color:#070910;background:linear-gradient(135deg,var(--gold),var(--gold2));border-color:transparent;
 box-shadow:0 4px 14px rgba(212,175,55,.35);}

/* ===== ALERT BANNER ===== */
.alertbar{position:relative;z-index:5;margin:0 18px 0;padding:11px 16px;border-radius:11px;
 display:flex;align-items:center;gap:11px;font-size:13px;font-weight:700;animation:slidein .5s ease;}
.alertbar.bad{background:rgba(255,92,108,.13);border:1px solid rgba(255,92,108,.42);color:#ffb3ba;}
.alertbar.good{background:rgba(57,217,138,.11);border:1px solid rgba(57,217,138,.38);color:#8ff0c0;}
.alertbar .x{margin-left:auto;cursor:pointer;opacity:.6;font-weight:900;}
.alertbar .x:hover{opacity:1;}
@keyframes slidein{from{opacity:0;transform:translateY(-8px);}to{opacity:1;transform:none;}}

.wrap{position:relative;z-index:2;max-width:1280px;margin:0 auto;padding:18px 18px 70px;}

/* ===== HERO VERDICT ===== */
.hero{background:var(--panel);border:1px solid var(--line);backdrop-filter:blur(16px);
 border-radius:18px;padding:24px;margin-top:16px;box-shadow:0 10px 34px rgba(0,0,0,.35);position:relative;overflow:hidden;}
.hero::after{content:'';position:absolute;top:0;left:0;right:0;height:3px;
 background:linear-gradient(90deg,transparent,var(--gold),var(--gold2),var(--gold),transparent);
 background-size:200% 100%;animation:sh 3s linear infinite;}
@keyframes sh{to{background-position:-200% 0;}}
.herotop{font-size:11px;letter-spacing:2px;color:var(--dim);text-transform:uppercase;font-weight:700;}
.verdict{font-size:56px;font-weight:900;letter-spacing:3px;line-height:1.05;margin-top:6px;}
.verdict.go{color:var(--ok);text-shadow:0 0 32px rgba(57,217,138,.35);}
.verdict.nogo{color:var(--bad);text-shadow:0 0 32px rgba(255,92,108,.35);}
.verdictsub{margin-top:8px;color:var(--dim);font-size:14px;max-width:760px;line-height:1.6;}
.herometa{display:flex;gap:22px;flex-wrap:wrap;margin-top:16px;font-size:12px;color:var(--dim);}
.herometa b{color:var(--txt);}

/* ===== KPI TILES ===== */
.kpis{display:flex;gap:14px;flex-wrap:wrap;margin-top:18px;}
.kpi{flex:1;min-width:158px;background:var(--panel);border:1px solid var(--line);border-radius:15px;padding:16px;transition:.22s;}
.kpi:hover{transform:translateY(-4px);box-shadow:0 0 26px rgba(212,175,55,.26);border-color:rgba(212,175,55,.4);}
.kpi .l{font-size:10px;color:var(--dim);text-transform:uppercase;letter-spacing:1.3px;font-weight:700;}
.kpi .v{font-size:22px;font-weight:900;margin-top:7px;word-break:break-word;}
.kpi .s{font-size:11px;color:var(--dim);margin-top:3px;}
.kpi .v.ok{color:var(--ok);} .kpi .v.warn{color:var(--warn);} .kpi .v.bad{color:var(--bad);}

/* ===== PANELS ===== */
.panel{background:var(--panel);border:1px solid var(--line);backdrop-filter:blur(16px);
 border-radius:16px;margin-top:18px;overflow:hidden;}
.phead{display:flex;align-items:center;gap:9px;padding:14px 18px;border-bottom:1px solid var(--line);cursor:pointer;}
.phead h3{font-size:13px;color:var(--gold);letter-spacing:1.2px;text-transform:uppercase;font-weight:800;}
.qmark{width:15px;height:15px;border-radius:50%;border:1px solid var(--line);color:var(--dim);
 font-size:9px;display:flex;align-items:center;justify-content:center;cursor:help;position:relative;}
.qmark:hover .tip{display:block;}
.tip{display:none;position:absolute;top:20px;left:0;background:#0d1119;border:1px solid var(--line);
 border-radius:8px;padding:9px 11px;font-size:11px;color:var(--dim);width:250px;z-index:20;line-height:1.5;font-weight:400;text-transform:none;letter-spacing:0;}
.pcount{font-size:10px;color:var(--dim);border:1px solid var(--line);border-radius:999px;padding:2px 8px;}
.pclose{margin-left:auto;color:var(--dim);cursor:pointer;font-weight:900;font-size:14px;}
.pclose:hover{color:var(--gold);}
.pbody{padding:16px 18px;}
.panel.collapsed .pbody{display:none;}
.unavail{color:var(--dim);font-size:13px;text-align:center;padding:26px;letter-spacing:1px;}

table{width:100%;border-collapse:collapse;font-size:13px;}
th,td{text-align:left;padding:11px 12px;border-bottom:1px solid var(--line);vertical-align:top;}
th{color:var(--gold);font-size:10px;text-transform:uppercase;letter-spacing:1.2px;}
tr{animation:fadeup .45s ease both;}
@keyframes fadeup{from{opacity:0;transform:translateY(6px);}to{opacity:1;transform:none;}}
.mono{font-family:ui-monospace,Menlo,monospace;font-size:11.5px;color:var(--dim);word-break:break-word;}
.pill{display:inline-block;padding:3px 10px;border-radius:999px;font-size:11px;font-weight:800;white-space:nowrap;}
.pill.ok{background:rgba(57,217,138,.15);color:var(--ok);border:1px solid rgba(57,217,138,.42);}
.pill.warn{background:rgba(240,180,41,.15);color:var(--warn);border:1px solid rgba(240,180,41,.42);}
.pill.bad{background:rgba(255,92,108,.15);color:var(--bad);border:1px solid rgba(255,92,108,.42);}

/* gauge */
.gauge{height:9px;border-radius:999px;background:rgba(255,255,255,.07);overflow:hidden;margin-top:10px;}
.gauge i{display:block;height:100%;border-radius:999px;background:linear-gradient(90deg,var(--ok),#8ff0c0);
 transition:width 1.1s cubic-bezier(.3,1,.4,1);}
.gauge.bad i{background:linear-gradient(90deg,var(--bad),#ff9aa4);}

/* copilot prompt */
.prompt{background:#05070c;border:1px solid var(--line);border-radius:11px;padding:14px;margin-top:12px;}
.prompt h4{font-size:12px;color:var(--gold2);margin-bottom:7px;}
.prompt p{font-size:12px;color:var(--dim);line-height:1.65;font-family:ui-monospace,Menlo,monospace;}
.copybtn{margin-top:9px;border:1px solid rgba(212,175,55,.4);background:rgba(212,175,55,.1);
 color:var(--gold);border-radius:7px;padding:5px 12px;font-size:11px;font-weight:700;cursor:pointer;transition:.2s;}
.copybtn:hover{background:rgba(212,175,55,.2);}
.copybtn.done{color:var(--ok);border-color:rgba(57,217,138,.5);background:rgba(57,217,138,.14);}

ol.steps{margin-left:20px;color:var(--dim);font-size:13.5px;line-height:1.9;}
ol.steps b{color:var(--txt);}

/* ===== COMMAND PALETTE ===== */
.overlay{display:none;position:fixed;inset:0;z-index:100;background:rgba(0,0,0,.62);backdrop-filter:blur(5px);}
.overlay.on{display:block;}
.palette{max-width:620px;margin:12vh auto 0;background:#0c1017;border:1px solid var(--line);
 border-radius:15px;overflow:hidden;box-shadow:0 24px 70px rgba(0,0,0,.6);}
.palette input{width:100%;border:none;outline:none;background:transparent;color:var(--txt);
 padding:16px 18px;font-size:15px;border-bottom:1px solid var(--line);}
.presults{max-height:52vh;overflow-y:auto;}
.pres{padding:11px 18px;border-bottom:1px solid var(--line);cursor:pointer;font-size:13px;}
.pres:hover{background:rgba(212,175,55,.09);}
.pres .t{color:var(--txt);font-weight:600;}
.pres .d{color:var(--dim);font-size:11.5px;margin-top:2px;}

.foot{margin-top:34px;text-align:center;color:#6b7686;font-size:11px;line-height:1.9;}
.foot b{color:var(--dim);}
@media print{.cmdbar,.overlay,.copybtn,.pclose{display:none!important;}body{padding-top:0;background:#fff;color:#000;}}
</style></head><body>
<div class="bgfx"></div><div class="particles" id="pt"></div>

<div class="cmdbar">
  <div class="cmdrow">
    <div class="brand">🛡️ <span class="b1">MAC-</span><span class="b2">HANDOFF</span></div>
    <span class="vtag" id="vtag">v3.0</span>
    <span class="live"><span class="dot"></span>LIVE</span>
    <span class="statebadge" id="stateBadge">—</span>
    <span class="modebadge" id="modeBadge">—</span>
    <div class="spacer"></div>
    <div class="searchbox" onclick="openPal()"><span>🔍 Search / Commands</span><span class="kbd">⌘K</span></div>
    <button class="iconbtn" onclick="toggleTheme()" title="Toggle theme">◐</button>
    <button class="iconbtn" onclick="location.reload()" title="Reload">⟳</button>
    <button class="iconbtn" onclick="exportJSON()" title="Export JSON">⤓ EXPORT</button>
    <button class="iconbtn" onclick="window.print()" title="Print / PDF">⚙ REPORT</button>
    <button class="iconbtn" onclick="fs()" title="Fullscreen">⛶</button>
    <span class="clock" id="clk">--:--:--</span>
  </div>
  <div class="cmdrow2">
    <span class="lbl">VIEW</span>
    <div class="chip on" data-v="overview" onclick="setView(this)">OVERVIEW</div>
    <div class="chip" data-v="gates" onclick="setView(this)">GATES</div>
    <div class="chip" data-v="evidence" onclick="setView(this)">EVIDENCE</div>
    <div class="chip" data-v="jamf" onclick="setView(this)">JAMF</div>
    <div class="chip" data-v="copilot" onclick="setView(this)">COPILOT</div>
    <div class="spacer"></div>
    <span class="lbl" id="uptime">SESSION --</span>
  </div>
</div>

<div id="alertHost"></div>

<div class="wrap">

  <div class="hero" data-view="overview">
    <div class="herotop">Cryptographic Hand-Off Decision</div>
    <div class="verdict" id="verdict">—</div>
    <div class="verdictsub" id="verdictSub"></div>
    <div class="gauge" id="gaugeWrap"><i id="gauge" style="width:0%"></i></div>
    <div class="herometa" id="heroMeta"></div>
  </div>

  <div class="kpis" data-view="overview" id="kpis"></div>

  <div class="panel" data-view="overview gates" id="pGates">
    <div class="phead" onclick="tog('pGates',event)">
      <h3>Cryptographic Gate Results</h3>
      <div class="qmark">?<div class="tip">Each gate must PASS before the Staging Admin can be removed. A single FAIL forces NO-GO and blocks deletion — this is what prevents an orphaned encrypted disk.</div></div>
      <span class="pcount" id="gateCount">0</span>
      <span class="pclose" onclick="tog('pGates',event)">✕</span>
    </div>
    <div class="pbody"><table id="gateTbl"><tr><th>Gate</th><th>Status</th><th>Why it matters</th></tr></table></div>
  </div>

  <div class="panel" data-view="evidence" id="pEvid">
    <div class="phead" onclick="tog('pEvid',event)">
      <h3>Raw Evidence &amp; Commands</h3>
      <div class="qmark">?<div class="tip">The exact command run for each gate and the raw output captured on-device. Retained for audit alongside handoff.log.</div></div>
      <span class="pcount" id="evCount">0</span>
      <span class="pclose" onclick="tog('pEvid',event)">✕</span>
    </div>
    <div class="pbody"><table id="evTbl"><tr><th>Gate</th><th>Command</th><th>Raw output</th></tr></table></div>
  </div>

  <div class="panel" data-view="overview jamf" id="pJamf">
    <div class="phead" onclick="tog('pJamf',event)">
      <h3>Final Verification — Jamf (View-Only)</h3>
      <div class="qmark">?<div class="tip">This device pushed fresh inventory via local jamf recon. View-only Jamf access is enough to confirm the key escrowed and the owner is correct.</div></div>
      <span class="pcount">4 steps</span>
      <span class="pclose" onclick="tog('pJamf',event)">✕</span>
    </div>
    <div class="pbody"><ol class="steps" id="jamfSteps"></ol></div>
  </div>

  <div class="panel" data-view="copilot" id="pCop">
    <div class="phead" onclick="tog('pCop',event)">
      <h3>🤖 Copilot Diagnostic Prompts</h3>
      <div class="qmark">?<div class="tip">Copy-paste ready prompts for root-cause analysis and ServiceNow ticket generation from this run's data.</div></div>
      <span class="pcount">2</span>
      <span class="pclose" onclick="tog('pCop',event)">✕</span>
    </div>
    <div class="pbody" id="copBody"></div>
  </div>

  <div class="foot" id="foot"></div>
</div>

<div class="overlay" id="ov" onclick="if(event.target.id==='ov')closePal()">
  <div class="palette">
    <input id="palIn" placeholder="Search gates, evidence, actions…" oninput="renderPal()">
    <div class="presults" id="palRes"></div>
  </div>
</div>

<script>
HTMLHEAD

# inject payload + the runtime JS
{
  printf 'const RUN = '
  printf '%s' "$PAYLOAD"
  printf ';\n'
} >> "$HTML"

cat >> "$HTML" <<'HTMLTAIL'
// ---------- particles ----------
(function(){var c=document.getElementById('pt');
 for(var i=0;i<24;i++){var s=document.createElement('span');var z=2+Math.random()*4;
 s.style.left=(Math.random()*100)+'%';s.style.width=s.style.height=z+'px';
 s.style.animationDuration=(9+Math.random()*14)+'s';s.style.animationDelay=(Math.random()*12)+'s';
 s.style.opacity=(0.12+Math.random()*0.35);c.appendChild(s);}})();

// ---------- clock + uptime ----------
var t0=Date.now();
function tick(){
 var d=new Date();
 var loc=d.toLocaleTimeString('en-US',{hour12:false});
 var utc=d.toUTCString().split(' ')[4];
 document.getElementById('clk').innerHTML=loc+' <small>LOCAL · '+utc+' UTC</small>';
 var s=Math.floor((Date.now()-t0)/1000);
 var m=Math.floor(s/60), r=s%60;
 document.getElementById('uptime').textContent='SESSION '+m+'m '+(r<10?'0':'')+r+'s';
}
setInterval(tick,1000);tick();

// ---------- header state ----------
var isGo = RUN.decision==='GO';
document.getElementById('vtag').textContent='v'+RUN.version;
var sb=document.getElementById('stateBadge');
sb.textContent=RUN.decision; sb.className='statebadge '+(isGo?'go':'nogo');
document.getElementById('modeBadge').textContent=RUN.mode;

// ---------- alert banner ----------
(function(){
 var host=document.getElementById('alertHost');
 var div=document.createElement('div');
 if(!isGo){
   div.className='alertbar bad';
   div.innerHTML='⛔ NO-GO — '+RUN.gatesFail+' cryptographic gate(s) failed. The Staging Admin was NOT removed. Resolve the failing gate(s) below and re-run.<span class="x" onclick="this.parentNode.remove()">✕</span>';
 }else{
   div.className='alertbar good';
   div.innerHTML='✅ GO — all cryptographic gates passed. Ownership transfer is safe; disk cannot be orphaned.<span class="x" onclick="this.parentNode.remove()">✕</span>';
 }
 host.appendChild(div);
})();

// ---------- hero ----------
var v=document.getElementById('verdict');
v.textContent=RUN.decision; v.className='verdict '+(isGo?'go':'nogo');
document.getElementById('verdictSub').textContent = isGo
 ? 'All cryptographic preconditions are satisfied. '+RUN.primaryUser+' holds a Secure Token, is a Volume Owner, and is FileVault-enabled. A surviving token holder ('+RUN.survivor+') remains, so removing the Staging Admin cannot orphan the encrypted volume.'
 : 'One or more preconditions failed. Deletion of the Staging Admin was blocked to prevent an orphaned encrypted volume. Review the failing gates, remediate, then re-run this tool.';
var pct = RUN.gatesTotal>0 ? Math.round((RUN.gatesPass/RUN.gatesTotal)*100) : 0;
document.getElementById('gaugeWrap').className='gauge'+(isGo?'':' bad');
setTimeout(function(){document.getElementById('gauge').style.width=pct+'%';},250);
document.getElementById('heroMeta').innerHTML=
 '<span>Host <b>'+RUN.host+'</b></span>'+
 '<span>Model <b>'+RUN.model+' · '+RUN.chip+'</b></span>'+
 '<span>Serial <b>'+RUN.serial+'</b></span>'+
 '<span>macOS <b>'+RUN.osver+' ('+RUN.osbuild+')</b></span>'+
 '<span>Generated <b>'+RUN.generated+'</b></span>'+
 '<span>Runtime <b>'+RUN.durationSec+'s</b></span>';

// ---------- KPIs ----------
function kpi(l,val,sub,cls){return '<div class="kpi"><div class="l">'+l+'</div><div class="v '+(cls||'')+'">'+val+'</div><div class="s">'+(sub||'')+'</div></div>';}
document.getElementById('kpis').innerHTML=
 kpi('Gates Passed',RUN.gatesPass+' / '+RUN.gatesTotal,pct+'% of preconditions',isGo?'ok':'bad')+
 kpi('Primary User',RUN.primaryUser,'device owner')+
 kpi('Staging Admin',RUN.stagingAdmin,RUN.deleteResult,RUN.deleteResult==='Deleted'?'ok':(RUN.deleteResult.indexOf('Refused')>-1||RUN.deleteResult.indexOf('failed')>-1?'bad':'warn'))+
 kpi('Surviving Token',RUN.survivor,'anti-orphan guarantee',RUN.survivor!=='none'?'ok':'warn')+
 kpi('Owner Auth',RUN.authMethod.split(' ')[0],RUN.authMethod)+
 kpi('Preboot',RUN.prebootResult,'APFS login screen',RUN.prebootResult==='updated'?'ok':'warn')+
 kpi('Jamf Recon',RUN.reconResult,'inventory push',RUN.reconResult==='complete'?'ok':'warn')+
 kpi('PRK',RUN.prkResult.split(' ')[0],RUN.prkResult);

// ---------- gates table ----------
function pcls(s){return s==='PASS'?'ok':(s==='WARN'?'warn':'bad');}
var gt=document.getElementById('gateTbl'),gh='<tr><th>Gate</th><th>Status</th><th>Why it matters</th></tr>';
RUN.gates.forEach(function(g,i){
 gh+='<tr style="animation-delay:'+(i*60)+'ms"><td><b>'+g.name+'</b></td><td><span class="pill '+pcls(g.status)+'">'+g.status+'</span></td><td style="color:var(--dim)">'+g.why+'</td></tr>';
});
gt.innerHTML=gh;
document.getElementById('gateCount').textContent=RUN.gates.length+' gates';

// ---------- evidence table ----------
var et=document.getElementById('evTbl'),eh='<tr><th>Gate</th><th>Command</th><th>Raw output</th></tr>';
RUN.gates.forEach(function(g,i){
 eh+='<tr style="animation-delay:'+(i*60)+'ms"><td><b>'+g.name+'</b></td><td class="mono">'+g.cmd+'</td><td class="mono">'+g.detail+'</td></tr>';
});
et.innerHTML=eh;
document.getElementById('evCount').textContent=RUN.gates.length+' records';

// ---------- jamf steps ----------
document.getElementById('jamfSteps').innerHTML=
 '<li>Open <b>Jamf Pro</b> and search by <b>username: '+RUN.primaryUser+'</b> or <b>serial: '+RUN.serial+'</b>.</li>'+
 '<li>Open the device record → <b>Inventory › Disk Encryption</b>.</li>'+
 '<li>Confirm the <b>Personal Recovery Key</b> is present and <b>FileVault 2 Enabled Users</b> lists <b>'+RUN.primaryUser+'</b>.</li>'+
 '<li>Confirm the record shows <b>'+RUN.primaryUser+'</b> as owner and the Staging Admin is gone.</li>';

// ---------- copilot prompts ----------
var failed=RUN.gates.filter(function(g){return g.status!=='PASS';}).map(function(g){return g.name+' ('+g.status+')';}).join('; ')||'none';
var P1='Analyze this macOS cryptographic hand-off telemetry. Decision='+RUN.decision+'. Host='+RUN.host+' Serial='+RUN.serial+' macOS='+RUN.osver+' ('+RUN.osbuild+'). Primary User='+RUN.primaryUser+', Staging Admin='+RUN.stagingAdmin+', surviving token holder='+RUN.survivor+'. Gates passed '+RUN.gatesPass+'/'+RUN.gatesTotal+'. Non-passing gates: '+failed+'. For each non-passing gate, explain the root cause on Apple Silicon, the risk of proceeding, and the safest remediation path (Secure Token grant, bootstrap token escrow, FileVault enablement) without orphaning the encrypted volume. Prioritize by risk.';
var P2='Using this macOS hand-off result (decision='+RUN.decision+', primary user='+RUN.primaryUser+', staging admin='+RUN.stagingAdmin+', result='+RUN.deleteResult+', serial='+RUN.serial+'), draft (A) a ServiceNow RITM/INFO ticket documenting the cryptographic ownership transfer with Executive Impact & Business Value, Description, Work Notes (investigation + resolution), and Status; and (B) a short white-glove Teams message to the device owner confirming completion. No IT jargon in the user message.';
document.getElementById('copBody').innerHTML=
 '<div class="prompt"><h4>📋 Root Cause / Safe Remediation</h4><p id="p1">'+P1+'</p><button class="copybtn" onclick="cp(\'p1\',this)">Copy prompt</button></div>'+
 '<div class="prompt"><h4>📋 ServiceNow / User Communication</h4><p id="p2">'+P2+'</p><button class="copybtn" onclick="cp(\'p2\',this)">Copy prompt</button></div>';
function cp(id,b){var t=document.getElementById(id).textContent;
 navigator.clipboard.writeText(t).then(function(){b.textContent='✓ Copied';b.classList.add('done');
 setTimeout(function(){b.textContent='Copy prompt';b.classList.remove('done');},2000);});}

// ---------- footer ----------
document.getElementById('foot').innerHTML=
 '<b>MAC-HANDOFF v'+RUN.version+'</b> — Cryptographic Hand-Off Console · Part of the MAC-MAINT toolkit<br>'+
 'Evidence captured on-device at '+RUN.generated+' · Log + payload retained at <b>'+RUN.outdir+'</b><br>'+
 'Gates are evaluated BEFORE any deletion. A single FAIL blocks removal of the Staging Admin — this is the anti-orphan guarantee.<br>'+
 'Intuitive confidential — For internal IT use only.';

// ---------- views ----------
function setView(el){
 document.querySelectorAll('.chip').forEach(function(c){c.classList.remove('on');});
 el.classList.add('on');
 var v=el.dataset.v;
 document.querySelectorAll('[data-view]').forEach(function(n){
   n.style.display = (v==='overview' ? (n.dataset.view.indexOf('overview')>-1?'':'none')
                                     : (n.dataset.view.indexOf(v)>-1?'':'none'));
 });
}
function tog(id,e){if(e)e.stopPropagation();document.getElementById(id).classList.toggle('collapsed');}
function toggleTheme(){document.body.classList.toggle('light');}
function fs(){if(!document.fullscreenElement){document.documentElement.requestFullscreen();}else{document.exitFullscreen();}}
function exportJSON(){
 var b=new Blob([JSON.stringify(RUN,null,2)],{type:'application/json'});
 var a=document.createElement('a');a.href=URL.createObjectURL(b);
 a.download='MAC-HANDOFF_'+RUN.serial+'_'+RUN.timestamp+'.json';a.click();
}

// ---------- command palette ----------
var CMDS=[
 {t:'Export run as JSON',d:'Download the full telemetry payload',a:exportJSON},
 {t:'Print / Save as PDF',d:'Generate an audit copy',a:function(){window.print();}},
 {t:'Toggle theme',d:'Dark ⇄ light console',a:toggleTheme},
 {t:'Fullscreen',d:'Expand the console',a:fs},
 {t:'View: Overview',d:'Verdict + KPIs',a:function(){setView(document.querySelector('[data-v=overview]'));}},
 {t:'View: Gates',d:'Cryptographic gate results',a:function(){setView(document.querySelector('[data-v=gates]'));}},
 {t:'View: Evidence',d:'Raw commands and output',a:function(){setView(document.querySelector('[data-v=evidence]'));}},
 {t:'View: Jamf',d:'View-only verification steps',a:function(){setView(document.querySelector('[data-v=jamf]'));}},
 {t:'View: Copilot',d:'Diagnostic prompts',a:function(){setView(document.querySelector('[data-v=copilot]'));}}
];
RUN.gates.forEach(function(g){CMDS.push({t:'Gate: '+g.name+' — '+g.status,d:g.why,a:function(){setView(document.querySelector('[data-v=gates]'));}});});
function openPal(){document.getElementById('ov').classList.add('on');document.getElementById('palIn').value='';renderPal();document.getElementById('palIn').focus();}
function closePal(){document.getElementById('ov').classList.remove('on');}
function renderPal(){
 var q=document.getElementById('palIn').value.toLowerCase();
 var h='';
 CMDS.filter(function(c){return !q||c.t.toLowerCase().indexOf(q)>-1||c.d.toLowerCase().indexOf(q)>-1;})
 .forEach(function(c,i){h+='<div class="pres" onclick="CMDS['+CMDS.indexOf(c)+'].a();closePal()"><div class="t">'+c.t+'</div><div class="d">'+c.d+'</div></div>';});
 document.getElementById('palRes').innerHTML=h||'<div class="pres"><div class="d">No matches.</div></div>';
}
document.addEventListener('keydown',function(e){
 if((e.metaKey||e.ctrlKey)&&e.key.toLowerCase()==='k'){e.preventDefault();openPal();}
 if(e.key==='Escape')closePal();
});
setView(document.querySelector('[data-v=overview]'));
</script></body></html>
HTMLTAIL

banner "COMPLETE"
ok "Console:  $HTML"
ok "Payload:  $JSONF"
ok "Log:      $LOG"
[ -n "${SUDO_USER:-}" ] && chown -R "$SUDO_USER" "$OUTDIR" 2>/dev/null
open "$HTML" 2>/dev/null

# ---- Phase 5 cold-boot reboot ------------------------------------------------
if [ "$DECISION" = "GO" ] && [ $AUDIT_ONLY -eq 0 ] && [ $DRY_RUN -eq 0 ] && [ $DO_REBOOT -eq 1 ] && [ "$DELETE_RESULT" = "Deleted" ]; then
  echo ""
  echo "${C_Y}Rebooting in 15s for Phase 5 cold-boot validation (Ctrl+C to cancel)...${C_0}"
  sleep 15
  priv reboot
fi

exit $( [ $GATES_FAILED -gt 0 ] && echo 1 || echo 0 )
