# PDF → DOCX Converter

[![Python](https://img.shields.io/badge/Python-3.9+-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-blue)](#)
[![License](https://img.shields.io/badge/License-MIT-green)](#)

A lightweight, cross-platform **desktop GUI** that converts PDF files into
editable **Word (.docx)** documents while faithfully preserving **images,
tables, paragraphs, and page layout**.

No file leaves your machine — everything is converted locally.

---

## Features

- 🖱️ **Drag & drop** a PDF onto the window, or pick one with **Browse…**
- 📄 Choose the output `.docx` location (defaults beside the source PDF)
- ⚙️ Converts in a **background thread** — the UI stays responsive
- 🖼️ **Images** and **tables** are preserved with layout fidelity
- 📦 Works on **Windows, macOS, and Linux**
- 🚀 Zero cloud dependency, fully offline

---

## Requirements

- **Python 3.9+** (3.12 recommended)
- `pip`

---

## Installation & Run

### Windows

Double-click **`run_windows.bat`**. It creates a virtual environment and
installs all dependencies automatically on first run.

### macOS / Linux

```bash
chmod +x run_macos.sh
./run_macos.sh
```

### Manual (any OS)

```bash
python -m venv .venv
source .venv/bin/activate            # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python pdf2docx_gui.py
```

> 💡 You can also pass a PDF as a command-line argument to preload it:
> `python pdf2docx_gui.py path/to/file.pdf`

---

## Usage

1. **Launch** the app.
2. **Drop** a PDF onto the window (or click **Browse…**).
3. Confirm the **output path** (defaults to `<source-name>.docx`).
4. Click **Convert**.
5. Open the `.docx` in Word / LibreOffice / Google Docs.

---

## Project layout

```
pdf2docx-gui/
├── pdf2docx_gui.py      # the application (Tkinter GUI + conversion)
├── requirements.txt     # dependencies
├── run_windows.bat      # Windows launcher (auto venv setup)
├── run_macos.sh         # macOS/Linux launcher (auto venv setup)
└── README.md            # this file
```

---

## How it works

The tool uses [`pdf2docx`](https://github.com/ArtifexSoftware/pdf2docx),
which is built on **PyMuPDF** (rendering & text extraction) and
**python-docx** (Word document generation). It parses each PDF page into
text blocks, vector shapes, and embedded images, then reassembles them
into a native `.docx` — keeping tables and figures as real Word objects
rather than flattened screenshots.

---

## Known limitations

- Complex vector graphics and exotic embedded fonts may render
  approximately (reflowed or substituted) rather than pixel-perfect.
- Very large, image-heavy PDFs can take a while; the progress bar shows
  activity but not a percentage.
- Scanned/image-only PDFs will not produce selectable text (OCR is not
  included).

---

## License

[MIT](./LICENSE)