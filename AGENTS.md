# Repository Guidelines

## Project Structure & Modules

This Python desktop app keeps its tracked code at the repository root. `pdf2docx_gui.py` starts the GUI and PDF CLI; `fluent_ui.py` holds the Fluent pages; `latex_to_docx.py` converts LaTeX with Pandoc. `snip_regions.py` crops PDF tables and equations, `polish.py` repairs converted documents, and `check.py` compares DOCX content with its PDF. Dependencies are in `requirements.txt`; `run_windows.bat` and `run_macos.sh` launch the app. Regression tests are in `tests/`.

## Build, Run & Validate

There is no build step. In a virtual environment, run `pip install -r requirements.txt`, then `python pdf2docx_gui.py`. Launchers create the environment and repair missing Fluent Widgets or Pandoc dependencies. PDF scripting uses `python pdf2docx_gui.py --cli input.pdf [output.docx] [--no-images] [--no-tidy]`. The LaTeX page writes a sibling `_from_tex.docx` and uses the source folder for figures and declared bibliographies. Check PDF output with `python check.py input.pdf output.docx`.

## Coding Style & Naming

Use Python with four spaces per indentation level. Follow the existing `snake_case` naming for modules, functions, and variables; use descriptive names for conversion and document-repair logic. Keep dependencies in `requirements.txt` and maintain compatibility with the Python versions stated in `README.md`. No formatter or linter is configured, so keep edits consistent with nearby code and review them before submitting.

## Testing Guidelines

Tests use the standard-library `unittest` runner; run `python -m unittest discover -s tests -v` (GUI tests use Qt's offscreen platform). Name new test modules `test_*.py`. For PDF changes, run the CLI and `check.py` on a representative document. For LaTeX changes, verify extracted text, editable tables/math, and embedded figures in the output DOCX. For GUI changes, include a brief manual check of the relevant interaction.

## Commits & Pull Requests

Recent commits use short Conventional Commit prefixes such as `feat:` and `docs:`. Use a focused message, for example `fix: preserve equation crops`. Pull requests should explain the user-visible change, note validation performed, and link a related issue when one exists. Include screenshots for visible GUI changes and a sample PDF/DOCX description for conversion changes; do not commit private source documents or generated outputs.

## Configuration & Local Files

Conversion runs locally and needs no credentials. Keep virtual environments and generated `.docx` outputs out of commits; `.gitignore` already excludes them.
