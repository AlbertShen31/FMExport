# Screen Data Scanner

A cross-platform Python application that captures screen regions, extracts tabular data using OCR, and exports it to CSV format. Works on both macOS and Windows.

## Features

- **Window Selection**: Choose a specific window to scan from a list of open windows
- **Screen Capture**: Select any area on your screen to capture manually
- **OCR Data Extraction**: Automatically extracts tabular data using Tesseract OCR
- **CSV Export**: Export extracted data to CSV format
- **Cross-Platform**: Works on both macOS and Windows
- **User-Friendly GUI**: Simple and intuitive interface

## Prerequisites

### 1. Python 3.8 or higher
Make sure Python is installed on your system.

### 2. uv (Python package installer)
Install `uv` for fast package management:

**macOS/Linux (Recommended):**
```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```
After installation, restart your terminal or run: `source ~/.cargo/env`

**Alternative methods for macOS:**

If the above doesn't work, try:
```bash
# Using pip (may require --break-system-packages on newer macOS)
python3 -m pip install --break-system-packages uv

# Or using Homebrew
brew install uv

# Or use the provided installation script
./install_uv.sh
```

**Windows:**
```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

**Note:** If `pip` command is broken (points to missing Python version), use `python3 -m pip` instead.

### 3. Tesseract OCR

**macOS:**
```bash
brew install tesseract
```

**Windows:**
1. Download Tesseract installer from: https://github.com/UB-Mannheim/tesseract/wiki
2. Install it (default location: `C:\Program Files\Tesseract-OCR`)
3. Add Tesseract to your PATH, or update the path in the code if needed

**Linux (Ubuntu/Debian):**
```bash
sudo apt-get install tesseract-ocr
```

## Installation

1. Clone or download this repository

2. Install Python dependencies using uv:
```bash
uv pip install -r requirements.txt
```

Or if you prefer to use uv with a virtual environment:
```bash
uv venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
uv pip install -r requirements.txt
```

## Usage

### Running from Source

```bash
python screen_scanner.py
```

### Using the Application

1. **Select Window or Area**: 
   - **Select Window**: Click "Select Window" button to choose from a list of open windows
     - A dialog will show all available windows
     - Select the window you want to scan
     - The window will be automatically captured (Windows) or you'll be prompted to select the area (macOS)
   - **Select Area**: Click "Select Area to Scan" button for manual selection
     - A full-screen overlay will appear
     - Click and drag to select the region containing your data
     - Press ESC to cancel

2. **Extract Data**: Click "Extract Data" button
   - The application will process the captured image using OCR
   - This may take a few moments depending on image size

3. **Export to CSV**: Click "Export to CSV" button
   - Choose a location and filename to save the CSV file
   - The data will be exported with a timestamp in the filename

## Building Executables

### For macOS

```bash
pyinstaller --onefile --windowed --name "ScreenDataScanner" --icon=NONE screen_scanner.py
```

The executable will be in the `dist` folder.

### For Windows

```bash
pyinstaller --onefile --windowed --name "ScreenDataScanner" --icon=NONE screen_scanner.py
```

**Note**: When building for Windows, you may need to:
- Include Tesseract OCR with your executable, or
- Ensure Tesseract is installed on the target machine

### Advanced Build Options

To include Tesseract OCR in the executable (if needed):

```bash
# macOS
pyinstaller --onefile --windowed --name "ScreenDataScanner" \
  --add-binary "/usr/local/bin/tesseract:tesseract" \
  screen_scanner.py

# Windows (adjust path as needed)
pyinstaller --onefile --windowed --name "ScreenDataScanner" \
  --add-binary "C:/Program Files/Tesseract-OCR/tesseract.exe;." \
  screen_scanner.py
```

## Troubleshooting

### Tkinter Not Found (macOS)
If you get `ModuleNotFoundError: No module named '_tkinter'`:

**Solution:**
```bash
brew install python-tk
```

This will install Python with tkinter support. You may need to recreate your virtual environment after installation:
```bash
# If using uv
uv venv
source .venv/bin/activate
uv pip install -r requirements.txt

# If using standard venv
python3 -m venv venv
source venv/bin/activate
uv pip install -r requirements.txt
```

### Tesseract Not Found
- Make sure Tesseract is installed and in your system PATH
- On Windows, you may need to specify the path explicitly in the code

### Poor OCR Results
- Ensure the selected area has clear, high-contrast text
- Try selecting a smaller, more focused area
- Make sure the text is not too small or blurry

### Import Errors
- Make sure all dependencies are installed: `uv pip install -r requirements.txt`
- Use a virtual environment to avoid conflicts
- If you're in a virtual environment and get tkinter errors, recreate the venv after installing python-tk

## Dependencies

- `mss`: Screen capture
- `Pillow`: Image processing
- `pytesseract`: OCR wrapper for Tesseract
- `pandas`: Data manipulation and CSV export
- `opencv-python`: Image preprocessing
- `numpy`: Numerical operations
- `pyinstaller`: Creating executables
- `tkinter`: GUI (usually included with Python)
- `pyobjc-framework-Quartz`: Window management on macOS (optional, for better window capture)
- `pywin32`: Window management on Windows (optional, for better window capture)

## License

This project is provided as-is for personal and commercial use.

## Notes

- The OCR accuracy depends on image quality and text clarity
- For best results, select areas with clear, well-contrasted text
- The application works best with structured tabular data
- Large images may take longer to process

