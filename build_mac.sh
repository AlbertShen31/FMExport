#!/bin/bash
# Build script for macOS executable

echo "Building macOS executable..."

# Check if PyInstaller is installed
if ! command -v pyinstaller &> /dev/null; then
    echo "PyInstaller not found. Installing..."
    uv pip install pyinstaller
fi

# Build the executable
pyinstaller --onefile \
    --windowed \
    --name "ScreenDataScanner" \
    --add-data "README.md:." \
    --hidden-import=pytesseract \
    --hidden-import=PIL \
    --hidden-import=cv2 \
    --hidden-import=pandas \
    --hidden-import=numpy \
    screen_scanner.py

echo "Build complete! Executable is in the 'dist' folder."
echo "Note: Users will still need Tesseract OCR installed on their system."

