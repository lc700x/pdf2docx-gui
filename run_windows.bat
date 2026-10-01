@echo off
rem Windows launcher
setlocal
cd /d "%~dp0"
if not exist .venv\Scripts\python.exe (
    echo Creating virtual environment...
    py -m venv .venv
    .venv\Scripts\python.exe -m pip install --upgrade pip
    .venv\Scripts\python.exe -m pip install -r requirements.txt
    if errorlevel 1 exit /b 1
)
.venv\Scripts\python.exe -c "import qfluentwidgets, pypandoc; pypandoc.get_pandoc_version()" >nul 2>&1
if errorlevel 1 (
    echo Installing application dependencies...
    .venv\Scripts\python.exe -m pip install -r requirements.txt
    if errorlevel 1 exit /b 1
)
call .venv\Scripts\python.exe pdf2docx_gui.py %*
