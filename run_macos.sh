#!/usr/bin/env bash
# macOS/Linux launcher
cd "$(dirname "$0")" || exit 1
if [ ! -d .venv ]; then
  echo "Creating virtual environment..."
  python3 -m venv .venv
  .venv/bin/python -m pip install --upgrade pip
  .venv/bin/python -m pip install -r requirements.txt || exit 1
fi
if ! .venv/bin/python -c "import qfluentwidgets, pypandoc; pypandoc.get_pandoc_version()" >/dev/null 2>&1; then
  echo "Installing application dependencies..."
  .venv/bin/python -m pip install -r requirements.txt || exit 1
fi
exec .venv/bin/python pdf2docx_gui.py "$@"
