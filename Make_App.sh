#!/bin/bash
# Creates a double-click "Tax Control Center" app (no Terminal window). Run once: bash Make_App.sh
cd "$(dirname "$0")" || exit 1
DIR="$(pwd)"
APP="$DIR/Tax Control Center.app"
if [ ! -x "$DIR/.venv/bin/python" ]; then echo "Folder .venv not found. Do the install step first."; exit 1; fi
rm -rf "$APP"
osacompile -o "$APP" -e "do shell script \"cd '$DIR' && ./.venv/bin/python control_center.py > /dev/null 2>&1 &\"" || exit 1
echo "Created: $APP"
echo "Drag it to your Dock or Desktop. Double-click to open."
