# Repository Guidelines

## Project Structure & Modules

This is a small Python desktop application; the tracked project files live at the repository root. `pdf2docx_gui.py` is the GUI and CLI entry point, and `fluent_ui.py` contains the PySide6 Fluent Widgets interface. `snip_regions.py` finds and crops tables and equations, while `polish.py` repairs document formatting after conversion. `check.py` checks converted DOCX content against its source PDF. Dependencies are listed in `requirements.txt`; `run_windows.bat` and `run_macos.sh` set up and launch the app. Focused regression tests live in `tests/`.

## Build, Run & Validate

There is no build step. Install dependencies in a virtual environment with `pip install -r requirements.txt`, then run `python pdf2docx_gui.py` to open the app. The launchers create the environment on first run and install dependencies if Fluent Widgets is missing. For scripting, use `python pdf2docx_gui.py --cli input.pdf [output.docx] [--no-images] [--no-tidy]`. Check a conversion with `python check.py input.pdf output.docx`; it exits nonzero when it finds missing text or spacing problems. Text inside cropped table or equation images is included through the image descriptions.

## Coding Style & Naming

Use Python with four spaces per indentation level. Follow the existing `snake_case` naming for modules, functions, and variables; use descriptive names for conversion and document-repair logic. Keep dependencies in `requirements.txt` and maintain compatibility with the Python versions stated in `README.md`. No formatter or linter is configured, so keep edits consistent with nearby code and review them before submitting.

## Testing Guidelines

Tests use the standard-library `unittest` runner; run `python -m unittest discover -s tests -v` (GUI tests use Qt's offscreen platform). Name new test modules `test_*.py`. For conversion changes, run the CLI on a representative PDF, then run `check.py` on the resulting DOCX and inspect it in Word or LibreOffice when layout is affected. For GUI changes, include a brief manual check of the relevant interaction.

## Commits & Pull Requests

Recent commits use short Conventional Commit prefixes such as `feat:` and `docs:`. Use a focused message, for example `fix: preserve equation crops`. Pull requests should explain the user-visible change, note validation performed, and link a related issue when one exists. Include screenshots for visible GUI changes and a sample PDF/DOCX description for conversion changes; do not commit private source documents or generated outputs.

## Configuration & Local Files

Conversion runs locally and needs no credentials. Keep virtual environments and generated `.docx` outputs out of commits; `.gitignore` already excludes them.
