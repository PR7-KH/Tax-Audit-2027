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

if [ -x /Applications/LibreOffice.app/Contents/MacOS/soffice ] || command -v soffice >/dev/null 2>&1 || command -v libreoffice >/dev/null 2>&1; then
    echo 'TOI PDF renderer detected.'
else
    echo 'For TOI PDF export, also install LibreOffice in Applications: https://www.libreoffice.org/download/download-libreoffice/'
fi
