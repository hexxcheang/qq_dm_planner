#!/bin/bash
# Double-click (macOS) or run `bash start_app.command` to start the Invoice & DO generator.
cd "$(dirname "$0")" || exit 1

if [ ! -x .venv/bin/python ]; then
    echo "First run: setting up, this takes a few minutes..."
    if ! python3 -m venv .venv; then
        echo "Python 3 was not found. Install it from https://www.python.org/downloads/"
        read -r -p "Press Enter to close." _
        exit 1
    fi
fi
if ! .venv/bin/python -m pip install --disable-pip-version-check -q -r requirements.txt; then
    echo "Installing the required packages failed. Check your internet connection and try again."
    read -r -p "Press Enter to close." _
    exit 1
fi

echo
echo "The app is starting and will open in your browser at http://localhost:8501"
echo "Keep this window open while you use the app. Close it to stop the app."
echo
OPENER=$(command -v open || command -v xdg-open)
[ -n "$OPENER" ] && (sleep 6; "$OPENER" http://localhost:8501 >/dev/null 2>&1) &
.venv/bin/python -m streamlit run app.py
