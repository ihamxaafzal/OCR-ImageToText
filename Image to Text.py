import os
import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import re

# Set Tesseract OCR path
pytesseract.pytesseract.tesseract_cmd = r"C:\\Users\\Hamza\\AppData\\Local\\Programs\\Tesseract-OCR\\tesseract.exe"

def preprocess_image(image):
    """Enhanced image preprocessing for better OCR results"""
    image = image.convert('L')  # Convert to grayscale
    image = image.filter(ImageFilter.MedianFilter(size=3))  # Reduce noise
    image = image.filter(ImageFilter.SHARPEN)

    # Increase contrast
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(2.5)

    # Binarization (thresholding) to remove background noise
    image = image.point(lambda x: 0 if x < 150 else 255)

    return image

def clean_text(text):
    """Cleans OCR-extracted text to remove gibberish"""
    text = text.replace("\n", " ")  # Replace line breaks
    text = re.sub(r'[^a-zA-Z0-9\s.,!?\'"-]', '', text)  # Remove random symbols
    text = re.sub(r'\s+', ' ', text).strip()  # Remove extra spaces
    text = re.sub(r'(?<!\w)[a-zA-Z](?!\w)', '', text)  # Remove isolated letters
    return text

def extract_text_from_images(folder_path, output_excel, progress_bar, progress_label):
    """Extract and clean text from images"""
    data = []
    files = [f for f in os.listdir(folder_path) if f.lower().endswith(('png', 'jpg', 'jpeg', 'tiff', 'bmp', 'gif'))]

    if not files:
        messagebox.showerror("Error", "No valid image files found")
        return

    for i, file_name in enumerate(files):
        file_path = os.path.join(folder_path, file_name)
        try:
            with Image.open(file_path) as img:
                processed_img = preprocess_image(img)
                
                # Using a better OCR mode
                extracted_text = pytesseract.image_to_string(processed_img, config="--psm 4 --oem 3")

                cleaned_text = clean_text(extracted_text)
                data.append([file_name, cleaned_text])

        except Exception as e:
            data.append([file_name, f"ERROR: {str(e)}"])

        progress = int((i + 1) / len(files) * 100)
        progress_bar['value'] = progress
        progress_label.config(text=f"Processing {progress}%: {file_name}")
        root.update_idletasks()

    if data:
        df = pd.DataFrame(data, columns=['File Name', 'Extracted Text'])
        df.to_excel(output_excel, index=False)
        messagebox.showinfo("Success", f"Data saved to {output_excel}")

    progress_label.config(text="OCR complete!")

# GUI Functions
def select_folder():
    folder_path = filedialog.askdirectory()
    folder_entry.delete(0, tk.END)
    folder_entry.insert(0, folder_path)

def select_output_file():
    output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    output_entry.delete(0, tk.END)
    output_entry.insert(0, output_file)

def run_script():
    folder_path = folder_entry.get()
    output_excel = output_entry.get()
    
    if not folder_path or not output_excel:
        messagebox.showerror("Error", "Please select both folder and output file")
        return
    
    progress_bar['value'] = 0
    progress_label.config(text="Starting...")
    extract_text_from_images(folder_path, output_excel, progress_bar, progress_label)

# GUI Setup
root = tk.Tk()
root.title("Advanced OCR Text Extractor")
root.geometry("650x450")

# GUI Components
tk.Label(root, text="Image Folder:").grid(row=0, column=0, padx=10, pady=10, sticky='e')
folder_entry = tk.Entry(root, width=50)
folder_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=select_folder).grid(row=0, column=2, padx=10)

tk.Label(root, text="Output Excel File:").grid(row=1, column=0, padx=10, pady=10, sticky='e')
output_entry = tk.Entry(root, width=50)
output_entry.grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=select_output_file).grid(row=1, column=2, padx=10)

tk.Button(root, text="Extract Text", command=run_script, bg='#4CAF50', fg='white').grid(row=2, column=1, pady=20)

progress_label = tk.Label(root, text="Ready")
progress_label.grid(row=3, column=1)
progress_bar = ttk.Progressbar(root, length=400)
progress_bar.grid(row=4, column=1)

root.mainloop()  # Keeps the GUI running
