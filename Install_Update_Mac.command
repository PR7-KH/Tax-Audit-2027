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

if [ -d "/Applications/Microsoft Excel.app" ]; then
    echo 'Microsoft Excel detected: recommended TOI PDF renderer.'
else
    echo 'TOI PDF export uses desktop Microsoft Excel by default. LibreOffice can be selected as an alternative.'
fi
echo 'For Khmer PDFs, install the workbook fonts in Font Book (for example Khmer OS Siemreap), then restart Excel.'
