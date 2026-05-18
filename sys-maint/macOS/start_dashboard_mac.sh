#!/bin/bash
# ============================================================
#  macOS Maintenance Dashboard — Fully Automated Launcher
#  Usage: bash start_dashboard_mac.sh
#  Everything installs itself — no manual steps required.
# ============================================================

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
LAUNCHER="$SCRIPT_DIR/launcher_mac.py"
PORT=9292

clear
echo ""
echo "  ============================================================"
echo "   macOS Maintenance Dashboard  |  Auto-Setup"
echo "  ============================================================"
echo ""

# ── Step 1: Python 3 — find it, or fully install it ──────────────────────────
echo "  [1/4] Checking for Python 3..."

PYTHON_BIN=""

for candidate in python3 python; do
    if command -v "$candidate" &>/dev/null; then
        VER=$("$candidate" --version 2>&1)
        if [[ "$VER" == *"Python 3"* ]]; then
            PYTHON_BIN="$candidate"
            echo "         Found: $VER"
            break
        fi
    fi
done

if [[ -z "$PYTHON_BIN" ]]; then
    echo "         Python 3 not found. Auto-installing now..."
    echo ""

    # ── Try Homebrew install of Python if brew already exists ────────────────
    if command -v brew &>/dev/null; then
        echo "         Homebrew already installed — using it to install Python 3..."
        brew install python3 --quiet
    else
        # ── Homebrew itself not found — install it silently first ────────────
        echo "         Homebrew not found. Installing Homebrew first..."
        echo "         (This needs your sudo password once — it won't ask again)"
        echo ""

        # Non-interactive Homebrew install
        NONINTERACTIVE=1 /bin/bash -c \
            "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

        if [[ $? -ne 0 ]]; then
            echo "  [FATAL] Homebrew installation failed."
            echo "          Check your internet connection and try again."
            read -p "  Press Enter to exit..."
            exit 1
        fi

        echo "         Homebrew installed."
        echo ""

        # Add Homebrew to PATH for this session
        # Apple Silicon Macs install to /opt/homebrew; Intel Macs to /usr/local
        if [[ -f /opt/homebrew/bin/brew ]]; then
            eval "$(/opt/homebrew/bin/brew shellenv)"
        elif [[ -f /usr/local/bin/brew ]]; then
            eval "$(/usr/local/bin/brew shellenv)"
        fi

        echo "         Installing Python 3 via Homebrew..."
        brew install python3 --quiet
    fi

    # Refresh PATH so newly installed python3 is visible
    if [[ -f /opt/homebrew/bin/brew ]]; then
        eval "$(/opt/homebrew/bin/brew shellenv)"
    elif [[ -f /usr/local/bin/brew ]]; then
        eval "$(/usr/local/bin/brew shellenv)"
    fi

    # Verify install succeeded
    for candidate in python3 python; do
        if command -v "$candidate" &>/dev/null; then
            VER=$("$candidate" --version 2>&1)
            if [[ "$VER" == *"Python 3"* ]]; then
                PYTHON_BIN="$candidate"
                echo "         Python 3 installed successfully: $VER"
                break
            fi
        fi
    done

    if [[ -z "$PYTHON_BIN" ]]; then
        echo ""
        echo "  [FATAL] Python 3 installation failed after all attempts."
        echo "          Manual fallback: https://www.python.org/downloads/macos/"
        read -p "  Press Enter to exit..."
        exit 1
    fi
fi

echo "  [1/4] Python 3 OK"
echo ""

# ── Step 2: Verify dashboard files ───────────────────────────────────────────
echo "  [2/4] Checking dashboard files..."

if [[ ! -f "$LAUNCHER" ]]; then
    echo "  [ERROR] launcher_mac.py not found in: $SCRIPT_DIR"
    echo "          Make sure all 3 files are in the same folder."
    read -p "  Press Enter to exit..."
    exit 1
fi

if [[ ! -f "$SCRIPT_DIR/dashboard_mac.html" ]]; then
    echo "  [ERROR] dashboard_mac.html not found in: $SCRIPT_DIR"
    read -p "  Press Enter to exit..."
    exit 1
fi

echo "  [2/4] All files present"
echo ""

# ── Step 3: Cache sudo password ───────────────────────────────────────────────
echo "  [3/4] Caching sudo credentials for SUDO commands..."
echo "         Enter your Mac login password below."
echo "         It is used only to pre-authorize sudo — not stored anywhere."
echo ""
sudo -v
if [[ $? -ne 0 ]]; then
    echo "  [WARNING] sudo auth failed or skipped."
    echo "            Commands marked SUDO in the dashboard may not work."
fi
echo ""

# ── Step 4: Launch ────────────────────────────────────────────────────────────
echo "  [4/4] Starting Maintenance Dashboard..."
echo ""
echo "  ============================================================"
echo "   Server  : http://localhost:$PORT"
echo "   Browser : Opening automatically..."
echo "   Stop    : Press Ctrl+C here, or use the Shutdown button"
echo "  ============================================================"
echo ""

# Keep sudo token alive every 50s in background so SUDO commands never hang
while true; do sudo -n true; sleep 50; kill -0 "$$" 2>/dev/null || exit; done &
SUDO_KEEP=$!

cleanup() {
    kill "$SUDO_KEEP" 2>/dev/null
    echo ""
    echo "  Server stopped. Port $PORT is now free."
    exit 0
}
trap cleanup INT TERM

"$PYTHON_BIN" "$LAUNCHER"

kill "$SUDO_KEEP" 2>/dev/null
