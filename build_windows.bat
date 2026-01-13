@echo off
REM Build script for Windows executable

echo Building Windows executable...

REM Check if PyInstaller is installed
uv pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo PyInstaller not found. Installing...
    uv pip install pyinstaller
)

REM Build the executable
pyinstaller --onefile ^
    --windowed ^
    --name "ScreenDataScanner" ^
    --add-data "README.md;." ^
    --hidden-import=pytesseract ^
    --hidden-import=PIL ^
    --hidden-import=cv2 ^
    --hidden-import=pandas ^
    --hidden-import=numpy ^
    screen_scanner.py

echo Build complete! Executable is in the 'dist' folder.
echo Note: Users will still need Tesseract OCR installed on their system.
pause

