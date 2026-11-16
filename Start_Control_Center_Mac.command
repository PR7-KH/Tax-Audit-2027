#!/bin/bash
cd "$(dirname "$0")" || exit 1
if [ -x ".venv/bin/python" ]; then PY=".venv/bin/python"; else PY="python3"; fi
"$PY" control_center.py
status=$?
if [ "$status" -ne 0 ]; then
  echo "Program exited with status $status."
  read -r -p 'Press Enter to close...'
fi
exit "$status"
