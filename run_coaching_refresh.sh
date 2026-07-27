#!/bin/bash
# Shell wrapper for launchd scheduling
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
/usr/bin/python3 "$SCRIPT_DIR/weekly_refresh.py"
