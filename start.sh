#!/bin/bash
# Business Card Scanner — Local Server Launcher (Mac / Linux)

echo ""
echo " ===================================================="
echo "   Business Card Scanner — Local Server Launcher"
echo " ===================================================="
echo ""

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Check Python
if ! command -v python3 &>/dev/null; then
    echo " [ERROR] python3 not found. Install Python 3 from https://python.org"
    exit 1
fi

echo " Starting server on http://localhost:8080"
echo ""

# Print local IP for phone
IP=$(ifconfig 2>/dev/null | grep 'inet ' | grep -v '127.0.0.1' | awk '{print $2}' | head -1)
if [ -n "$IP" ]; then
    echo " From your phone (same WiFi): http://$IP:8080"
fi
echo ""
echo " Press Ctrl+C to stop the server when done."
echo ""

# Open browser
(sleep 1 && open "http://localhost:8080" 2>/dev/null || xdg-open "http://localhost:8080" 2>/dev/null) &

# Start server
cd "$SCRIPT_DIR" && python3 -m http.server 8080
