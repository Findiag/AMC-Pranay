#!/usr/bin/env bash
set -e
cd "$(dirname "${BASH_SOURCE[0]}")"
echo ""
echo "  ASK MY CFO — M1 Automation (Flask)"
echo "  ===================================="
echo ""
pip install -r requirements.txt --break-system-packages -q 2>/dev/null || pip install -r requirements.txt -q
echo "  → http://localhost:5000"
echo ""
python app.py
