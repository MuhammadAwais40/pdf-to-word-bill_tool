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
  --collect-all pdfminer ^
  --collect-all pypdfium2 ^
  --collect-all pypdfium2_raw ^
  --collect-data docx ^
  --collect-all webview ^
  main.py

echo.
echo Done! EXE yahan hai:  dist\PDF Converter.exe
pause
