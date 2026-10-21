@echo off
setlocal
cd /d "%~dp0"

echo [1/4] Virtual environment...
if not exist venv (
    python -m venv venv
)
call venv\Scripts\activate

echo [2/4] Installing requirements...
python -m pip install --upgrade pip
pip install -r requirements.txt

echo [3/4] Cleaning old build...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [4/4] Building EXE...
pyinstaller --noconfirm --clean --onefile --noconsole ^
  --name "PDF Converter" ^
  --icon "convert.ico" ^
  --add-data "templates;templates" ^
  --add-data "static;static" ^
  --add-data "convert.ico;." ^
  --collect-all pdfplumber ^
  --collect-data pdfminer ^
  --collect-all pypdfium2 ^
  --collect-all pypdfium2_raw ^
  --collect-data docx ^
  --collect-all webview ^
  --exclude-module tkinter ^
  --exclude-module numpy ^
  --exclude-module pandas ^
  --exclude-module scipy ^
  --exclude-module matplotlib ^
  --exclude-module PyQt5 ^
  --exclude-module PyQt6 ^
  --exclude-module PySide2 ^
  --exclude-module PySide6 ^
  --exclude-module qtpy ^
  --exclude-module gi ^
  --exclude-module IPython ^
  --exclude-module pytest ^
  main.py

echo.
echo Done! EXE location:  dist\PDF Converter.exe
pause
