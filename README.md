# Luma OCR — Image to Text

A privacy-friendly desktop app for extracting text from images with Tesseract OCR. Add individual images or a complete folder, review each image, tune recognition settings, then copy or export the results.

## Features

- Batch processing for PNG, JPEG, TIFF, BMP, GIF, and WebP images.
- Dark desktop interface with image preview, progress updates, and per-image OCR confidence estimates.
- Automatic comparison of original and enhanced image passes, with alternate page layouts for block and auto-page recognition.
- Optional contrast enhancement and noise reduction.
- Configurable Tesseract language and page segmentation mode.
- Copy results to the clipboard or export them as TXT, CSV, or Excel (.xlsx).
- Local processing: images and extracted text are processed on your computer.
- Automatic Tesseract detection, with an in-app option to choose its executable.

## Requirements

- Windows 10 or 11 (the supplied setup and launcher scripts are for Windows).
- Python 3.9 or newer.
- Tesseract OCR installed separately from the Python packages. Tesseract 5 is recommended.
- Tesseract trained data for each language you want to recognize.

tkinter is included with most standard Windows Python installers. It is part of Python's standard library and should not be installed with pip.

## Quick install on Windows

1. Download or clone this repository and open its folder.
2. Double-click setup.bat. It installs the Python packages and checks for Tesseract. If Python or Tesseract is missing, it attempts installation with Windows Package Manager (winget).
3. If Windows asks you to approve an installer, allow it to finish. Setup confirms that Tesseract is available before it reports success.
4. Double-click run.bat to start Luma OCR.

winget is included with current Windows 10/11 installations. If it is unavailable or blocked by your device policy, follow the manual setup below.

## Manual installation

### 1. Install Python

Install Python 3.9 or later from https://www.python.org/downloads/. On Windows, enable the Python launcher and the option to add Python to PATH if offered.

### 2. Install Tesseract OCR

Install Tesseract OCR for Windows using the UB Mannheim Tesseract installer: https://github.com/UB-Mannheim/tesseract/wiki. The common install location is:

    C:\Program Files\Tesseract-OCR\tesseract.exe

The app checks PATH and common installation folders. If it does not find Tesseract, choose tesseract.exe using the Choose Tesseract executable… option in the app.

### 3. Install Python dependencies

Open a terminal in the repository folder and run:

    python -m pip install -r requirements.txt

### 4. Launch the app

Double-click run.bat, or run:

    python "Image to Text.py"

## How to use

1. Select Add images to choose files, or Folder to add supported images from one folder.
2. Select an image in the queue to preview it.
3. Set the OCR language and page layout. The default language is eng and the default layout is Single block (6).
4. Optionally enable Improve contrast & remove noise for difficult scans.
5. Select Extract text and review the results in the Extracted text tab.
6. Select Copy text to copy the displayed results, or Export… to save a TXT, CSV, or Excel workbook.

The CSV and Excel exports include the filename, confidence estimate, extracted text, and any processing error for each image.

## Languages

Enter Tesseract language codes in the Language field. Examples include eng for English, urd for Urdu, and eng+urd for both. The corresponding trained-data files must be installed in Tesseract's tessdata directory. If a language is unavailable, Tesseract reports an OCR error; installing a Python package will not add language data.

## Troubleshooting

### “Tesseract not found” or “tesseract is not installed”

Install Tesseract OCR, restart the app, or use Choose Tesseract executable… to select tesseract.exe directly.

### A language is missing

Install that language's Tesseract trained data and confirm that its code is entered correctly. Multiple languages can be separated with +, for example eng+urd.

### Text is incomplete or inaccurate

OCR depends on image sharpness, resolution, contrast, orientation, and language. Try a clearer or higher-resolution image, enable preprocessing, or choose a page layout that matches the image. Review the extracted result before using it as a final transcript.

### Excel export is unavailable

Install the dependencies again with python -m pip install -r requirements.txt. Excel export uses pandas and openpyxl; TXT and CSV export use Python's standard library.

## Repository contents

    .
    ├── Image to Text.py   # Desktop application
    ├── requirements.txt   # Python dependencies
    ├── setup.bat          # Windows dependency setup
    ├── run.bat            # Windows app launcher
    ├── README.md          # Documentation
    ├── LICENSE            # MIT license
    └── .gitignore

## Contributing

Bug reports and focused improvements are welcome. Please include the Windows version, Python version, Tesseract version, language code, and a brief description of the issue. Do not include private documents or images in public issue reports.

## License

This project is licensed under the MIT License. See LICENSE for details.
