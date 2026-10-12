#!/usr/bin/env bash
# Runner script for Bosta Smart Order Manager
cd "$(dirname "$0")"
source .venv/bin/activate
export QT_LOGGING_RULES="qt.text.font.db=false;*.debug=false"
python3 app.py "$@"
