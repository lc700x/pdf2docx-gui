#!/usr/bin/env bash
# macOS/Linux launcher
cd "$(dirname "$0")" || exit 1
if [ ! -d .venv ]; then
  echo "Creating virtual environment..."
  python3 -m venv .venv
  source .venv/bin/activate
  pip install --upgrade pip
  pip install -r requirements.txt
fi
source .venv/bin/activate
python pdf2docx_gui.py "$@"