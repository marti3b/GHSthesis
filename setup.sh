#!/bin/bash
# Setup script — run this once before using analyze.py
set -e

echo ""
echo "══════════════════════════════════════════════════════"
echo "       Congruence Analysis — First-Time Setup"
echo "══════════════════════════════════════════════════════"
echo ""

# ── 1. Check for Homebrew ──────────────────────────────────────────────────────
if ! command -v brew &>/dev/null; then
    echo "Installing Homebrew (package manager for macOS)..."
    /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
    # add brew to path for Apple Silicon Macs
    if [ -f /opt/homebrew/bin/brew ]; then
        eval "$(/opt/homebrew/bin/brew shellenv)"
        echo 'eval "$(/opt/homebrew/bin/brew shellenv)"' >> ~/.zprofile
    fi
else
    echo "✓ Homebrew already installed"
fi

# ── 2. Check for Python 3 ─────────────────────────────────────────────────────
if ! command -v python3 &>/dev/null || python3 -c "import sys; sys.exit(0 if sys.version_info >= (3,9) else 1)" 2>/dev/null; then
    echo "Installing Python 3..."
    brew install python3
else
    echo "✓ Python 3 already installed ($(python3 --version))"
fi

# ── 3. Install Python packages ────────────────────────────────────────────────
echo ""
echo "Installing required Python packages..."
python3 -m pip install --upgrade pip --quiet
python3 -m pip install pdfplumber pandas openpyxl --quiet

echo ""
echo "✓ All packages installed"

# ── 4. Move example PDF to documents/ ────────────────────────────────────────
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
for pdf in "$SCRIPT_DIR"/*.pdf; do
    [ -f "$pdf" ] || continue
    dest="$SCRIPT_DIR/documents/$(basename "$pdf")"
    if [ ! -f "$dest" ]; then
        mv "$pdf" "$dest"
        echo "✓ Moved $(basename "$pdf") → documents/"
    fi
done

echo ""
echo "══════════════════════════════════════════════════════"
echo "  Setup complete! To run the analysis:"
echo ""
echo "  1. Drop PDF files into:  documents/"
echo "  2. Edit keywords in:     keywords.txt"
echo "  3. Run:                  python3 analyze.py"
echo "══════════════════════════════════════════════════════"
echo ""
