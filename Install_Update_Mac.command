#!/bin/bash
# Run with: bash Install_Update_Mac.command
set -euo pipefail
cd "$(dirname "$0")"
if [ ! -x .venv/bin/python ]; then
    python3 -m venv .venv
fi
./.venv/bin/python -m pip install -r requirements.txt
./.venv/bin/python -m pip check
echo 'Dependencies ready. Open Start_Control_Center_Mac.command to start the app.'
