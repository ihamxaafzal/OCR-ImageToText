# Advanced OCR Text Extractor

## Description
This project is a GUI-based tool for extracting text from images using OCR (Optical Character Recognition). It utilizes Python libraries such as `pytesseract`, `PIL`, `pandas`, and `tkinter` to process images, enhance them for better OCR results, and export extracted text into an Excel file.

## Features
- Supports multiple image formats (`png`, `jpg`, `jpeg`,....).
- Enhances images using filters to improve OCR accuracy.
- Extracts text using `Tesseract OCR` with optimized settings.
- Cleans extracted text to remove unnecessary characters.
- Provides a simple GUI for selecting input folders and output files.
- Displays real-time progress updates during text extraction.

## Prerequisites
Ensure you have Python installed on your system. You also need to install Tesseract OCR separately.

### Installing Tesseract OCR
1. Download and install Tesseract OCR from [here](https://github.com/UB-Mannheim/tesseract/wiki).
2. Locate the installation path (usually `C:\Program Files\Tesseract-OCR\tesseract.exe`).
3. Update the script to use your installation path:
   ```python
   pytesseract.pytesseract.tesseract_cmd = r"C:\\Users\\(your username)\\AppData\\Local\\Programs\\Tesseract-OCR\\tesseract.exe"
   ```
Your Location can be different. Just copy and paste your "tesseract.exe" location.

## Installation
To install the required Python libraries, run:
```bash
pip install pytesseract pillow pandas tkinter
```

## Usage
1. Run the script:
   ```bash
   python script.py
   ```
2. Select the image folder.
3. Select the output Excel file.
4. Click the `Extract Text` button.
5. View the extracted text in the saved Excel file.

## Dependencies
The script uses the following Python libraries:
- `os`
- `pytesseract`
- `PIL (Pillow)`
- `pandas`
- `tkinter`
- `re`
- 
## Install all required libraries.

## How It Works
1. The script loads and preprocesses images (grayscale conversion, noise reduction, contrast enhancement, thresholding).
2. Extracts text using `pytesseract.image_to_string()`.
3. Cleans the extracted text to remove unnecessary symbols and noise.
4. Saves the extracted text into an Excel file.

## License
This project is licensed under the MIT License.


