#!/usr/bin/env bash
# ==============================================================================
# Instagram Comment Exporter GUI Launcher
#
# Launches the local Streamlit dashboard on http://localhost:8501
# Usage:
#   ./run_gui.sh
#   ./run_gui.sh --port 8502
# ==============================================================================

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

PORT="${1:-8501}"
if [ "$1" == "--port" ] && [ -n "$2" ]; then
    PORT="$2"
fi

echo "================================================================="
echo "  Instagram Comment & Reply Exporter - Local Dashboard"
echo "================================================================="
echo "  Project Path: $SCRIPT_DIR"
echo "  Local Port:   $PORT"
echo "================================================================="

# Check Python 3
if ! command -v python3 &>/dev/null; then
    echo "[!] Error: python3 is not installed or not in PATH."
    exit 1
fi

# Check Streamlit
if ! python3 -c "import streamlit" &>/dev/null; then
    echo "[!] Streamlit is not installed. Installing required dependencies..."
    python3 -m pip install --user --break-system-packages streamlit openpyxl google-api-python-client google-auth
fi

echo "[+] Starting local dashboard..."
echo "[+] URL: http://localhost:$PORT"
echo "[+] Press Ctrl+C in this terminal to stop the GUI server."
echo ""

exec python3 -m streamlit run "$SCRIPT_DIR/gui_app.py" \
    --server.port "$PORT" \
    --server.headless true \
    --browser.gatherUsageStats false
