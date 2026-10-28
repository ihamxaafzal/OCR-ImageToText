import csv
import os
import shutil
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageEnhance, ImageFilter, ImageOps, ImageTk
import pytesseract
from pytesseract import Output

APP_TITLE = "Luma OCR | Image to Text"
SUPPORTED = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".gif", ".webp"}


def find_tesseract():
    configured = os.environ.get("TESSERACT_CMD")
    if configured and Path(configured).is_file():
        return configured
    found = shutil.which("tesseract")
    if found:
        return found
    for candidate in (r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                      str(Path.home() / r"AppData\Local\Programs\Tesseract-OCR\tesseract.exe")):
        if Path(candidate).is_file():
            return candidate
    return None


class LumaOCR:
    BG = "#080D18"
    CARD = "#111A2B"
    INK = "#E8ECF5"
    MUTED = "#9AA8BD"
    ACCENT = "#8878FF"
    GREEN = "#36D399"

    def __init__(self, root):
        self.root = root
        root.title(APP_TITLE)
        screen_w = root.winfo_screenwidth()
        screen_h = root.winfo_screenheight()
        window_w = min(1240, max(900, screen_w - 80))
        window_h = min(820, max(620, screen_h - 90))
        root.geometry(f"{window_w}x{window_h}")
        root.minsize(min(900, window_w), min(620, window_h))
        root.configure(bg=self.BG)
        self.paths = []
        self.results = []
        self.busy = False
        self.photo = None
        self._style()
        self._build()
        self._set_tesseract_status()

    def _style(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TFrame", background=self.BG)
        style.configure("Card.TFrame", background=self.CARD)
        style.configure("Hero.TFrame", background="#172238")
        style.configure("TLabel", background=self.BG, foreground=self.INK, font=("Segoe UI", 10))
        style.configure("Muted.TLabel", background=self.BG, foreground=self.MUTED, font=("Segoe UI", 9))
        style.configure("CardMuted.TLabel", background=self.CARD, foreground=self.MUTED, font=("Segoe UI", 9))
        style.configure("Title.TLabel", background=self.BG, foreground=self.INK, font=("Segoe UI", 25, "bold"))
        style.configure("Sub.TLabel", background=self.BG, foreground=self.MUTED, font=("Segoe UI", 10))
        style.configure("CardTitle.TLabel", background=self.CARD, foreground=self.INK, font=("Segoe UI", 13, "bold"))
        style.configure("HeroTitle.TLabel", background="#172238", foreground=self.INK, font=("Segoe UI", 17, "bold"))
        style.configure("HeroSub.TLabel", background="#172238", foreground="#AAB7CE", font=("Segoe UI", 10))
        style.configure("Badge.TLabel", background="#202D43", foreground="#BFC9DA", font=("Segoe UI", 9, "bold"), padding=(11, 7))
        style.configure("Ready.TLabel", background="#17372F", foreground="#6EE7B7", font=("Segoe UI", 9, "bold"), padding=(11, 7))
        style.configure("Missing.TLabel", background="#3A2A1D", foreground="#FDBA74", font=("Segoe UI", 9, "bold"), padding=(11, 7))
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), padding=(16, 11), background=self.ACCENT, foreground="white")
        style.map("Accent.TButton", background=[("active", "#6253EA"), ("disabled", "#35334F")])
        style.configure("Soft.TButton", font=("Segoe UI", 9), padding=(10, 7), background="#202A3D", foreground=self.INK)
        style.map("Soft.TButton", background=[("active", "#2B3851")])
        style.configure("TEntry", fieldbackground="#1B2538", foreground=self.INK, insertcolor=self.INK, bordercolor="#35425A")
        style.configure("TCombobox", fieldbackground="#1B2538", background="#202A3D", foreground=self.INK, arrowcolor=self.INK)
        style.map("TCombobox", fieldbackground=[("readonly", "#1B2538")], foreground=[("readonly", self.INK)])
        style.configure("TCheckbutton", background=self.CARD, foreground=self.INK)
        style.map("TCheckbutton", background=[("active", self.CARD)], foreground=[("disabled", self.MUTED)])
        style.configure("TProgressbar", thickness=6, troughcolor="#202A3D", background=self.ACCENT)
        style.configure("Dark.Vertical.TScrollbar", background="#202A3D", troughcolor=self.CARD, arrowcolor=self.MUTED, bordercolor=self.CARD)
        style.configure("Dark.Horizontal.TScrollbar", background="#202A3D", troughcolor=self.CARD, arrowcolor=self.MUTED, bordercolor=self.CARD)
        style.configure("TNotebook", background=self.CARD, borderwidth=0)
        style.configure("TNotebook.Tab", padding=(13, 8), font=("Segoe UI", 9), background="#1B2538", foreground=self.MUTED)
        style.map("TNotebook.Tab", background=[("selected", self.CARD), ("active", "#253149")], foreground=[("selected", self.INK)])
        style.configure("Treeview", rowheight=29, font=("Segoe UI", 9), background="#121A2A", fieldbackground="#121A2A", foreground=self.INK)
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), background="#202A3D", foreground=self.INK)

    def _build(self):
        outer = ttk.Frame(self.root, padding=(24, 13, 24, 12))
        outer.pack(fill="both", expand=True)

        header = ttk.Frame(outer)
        header.pack(fill="x", pady=(0, 11))
        mark = tk.Canvas(header, width=48, height=48, bg=self.BG, highlightthickness=0)
        mark.pack(side="left", padx=(0, 12))
        mark.create_oval(2, 2, 46, 46, fill=self.ACCENT, outline="")
        mark.create_text(24, 23, text="L", fill="white", font=("Segoe UI", 21, "bold"))
        brand = ttk.Frame(header)
        brand.pack(side="left", anchor="center")
        ttk.Label(brand, text="LUMA OCR", style="Title.TLabel").pack(anchor="w")
        ttk.Label(brand, text="IMAGE TO TEXT WORKSPACE", style="Sub.TLabel").pack(anchor="w", pady=(1, 0))
        self.engine_badge = ttk.Label(header, text="●  CHECKING ENGINE", style="Badge.TLabel")
        self.engine_badge.pack(side="right", anchor="center")

        hero = ttk.Frame(outer, style="Hero.TFrame", padding=(18, 10))
        hero.pack(fill="x", pady=(0, 12))
        accent = tk.Canvas(hero, width=5, height=48, bg="#172238", highlightthickness=0)
        accent.pack(side="left", padx=(0, 15))
        accent.create_line(2, 2, 2, 46, fill=self.ACCENT, width=4, capstyle="round")
        hero_copy = ttk.Frame(hero, style="Hero.TFrame")
        hero_copy.pack(side="left", fill="x", expand=True)
        ttk.Label(hero_copy, text="Turn pictures into editable text.", style="HeroTitle.TLabel").pack(anchor="w")
        ttk.Label(hero_copy, text="Add a single image or process a full batch. Your files stay on this computer.", style="HeroSub.TLabel").pack(anchor="w", pady=(4, 0))
        ttk.Label(hero, text="PRIVATE  ·  BATCH READY", style="Badge.TLabel").pack(side="right", padx=(12, 0))

        self.main = ttk.Frame(outer)
        self.main.pack(fill="both", expand=True)
        self.main.columnconfigure(0, weight=0, minsize=330)
        self.main.columnconfigure(1, weight=1)
        self.main.rowconfigure(0, weight=1)
        self._left_panel()
        self._right_panel()

        footer = ttk.Frame(outer)
        footer.pack(fill="x", pady=(10, 0))
        self.status = ttk.Label(footer, text="Ready when you are", style="Muted.TLabel")
        self.status.pack(side="left")
        self.progress = ttk.Progressbar(footer, mode="determinate", maximum=100, length=200)
        self.progress.pack(side="right", fill="x", expand=True, padx=(24, 0))

    def _left_panel(self):
        shell = ttk.Frame(self.main, style="Card.TFrame")
        shell.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        shell.rowconfigure(0, weight=1)
        shell.columnconfigure(0, weight=1)
        sidebar_canvas = tk.Canvas(shell, bg=self.CARD, highlightthickness=0, borderwidth=0)
        sidebar_canvas.grid(row=0, column=0, sticky="nsew")
        sidebar_scroll = ttk.Scrollbar(shell, orient="vertical", command=sidebar_canvas.yview, style="Dark.Vertical.TScrollbar")
        sidebar_scroll.grid(row=0, column=1, sticky="ns")
        sidebar_canvas.configure(yscrollcommand=sidebar_scroll.set)
        card = ttk.Frame(sidebar_canvas, style="Card.TFrame", padding=17)
        window_id = sidebar_canvas.create_window((0, 0), window=card, anchor="nw")
        card.bind("<Configure>", lambda _event: sidebar_canvas.configure(scrollregion=sidebar_canvas.bbox("all")))
        sidebar_canvas.bind("<Configure>", lambda event: sidebar_canvas.itemconfigure(window_id, width=event.width))

        ttk.Label(card, text="Your images", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(card, text="Add one image or a whole folder.", style="CardMuted.TLabel").pack(anchor="w", pady=(3, 10))
        row = ttk.Frame(card, style="Card.TFrame")
        row.pack(fill="x")
        ttk.Button(row, text="＋  Add images", style="Soft.TButton", command=self.add_images).pack(side="left")
        ttk.Button(row, text="＋  Folder", style="Soft.TButton", command=self.add_folder).pack(side="left", padx=(7, 0))
        queue = ttk.Frame(card, style="Card.TFrame")
        queue.pack(fill="x", pady=(10, 6))
        queue.columnconfigure(0, weight=1)
        self.file_list = tk.Listbox(queue, height=4, relief="flat", bg="#0D1424", fg=self.INK,
                                    selectbackground="#443A7A", selectforeground="white", highlightthickness=0,
                                    font=("Segoe UI", 9), yscrollcommand=lambda first, last: queue_scroll.set(first, last),
                                    xscrollcommand=lambda first, last: queue_xscroll.set(first, last))
        self.file_list.grid(row=0, column=0, sticky="ew")
        queue_scroll = ttk.Scrollbar(queue, orient="vertical", command=self.file_list.yview, style="Dark.Vertical.TScrollbar")
        queue_scroll.grid(row=0, column=1, sticky="ns")
        queue_xscroll = ttk.Scrollbar(queue, orient="horizontal", command=self.file_list.xview, style="Dark.Horizontal.TScrollbar")
        queue_xscroll.grid(row=1, column=0, sticky="ew")
        ttk.Button(card, text="Remove selected", style="Soft.TButton", command=self.remove_selected).pack(anchor="w")

        ttk.Separator(card).pack(fill="x", pady=12)
        ttk.Label(card, text="OCR settings", style="CardTitle.TLabel").pack(anchor="w", pady=(0, 8))
        ttk.Label(card, text="Language (Tesseract codes)", style="CardMuted.TLabel").pack(anchor="w")
        self.lang = tk.StringVar(value="eng")
        ttk.Entry(card, textvariable=self.lang).pack(fill="x", pady=(4, 8))
        ttk.Label(card, text="Page layout", style="CardMuted.TLabel").pack(anchor="w")
        self.psm = tk.StringVar(value="Single block (6)")
        ttk.Combobox(card, textvariable=self.psm, state="readonly", values=["Single block (6)", "Auto page (3)", "Single line (7)", "Sparse text (11)"]).pack(fill="x", pady=(4, 7))
        self.preprocess = tk.BooleanVar(value=False)
        ttk.Checkbutton(card, text="Improve contrast & remove noise", variable=self.preprocess).pack(anchor="w")
        self.status_tess = ttk.Label(card, text="Checking Tesseract…", style="CardMuted.TLabel")
        self.status_tess.pack(anchor="w", pady=(8, 0))
        ttk.Button(card, text="Choose Tesseract executable…", style="Soft.TButton", command=self.choose_tesseract).pack(anchor="w", pady=(6, 0))

    def _right_panel(self):
        card = ttk.Frame(self.main, style="Card.TFrame", padding=18)
        card.grid(row=0, column=1, sticky="nsew")
        card.rowconfigure(1, weight=1)
        card.columnconfigure(0, weight=1)
        top = ttk.Frame(card, style="Card.TFrame")
        top.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        ttk.Label(top, text="Workspace", style="CardTitle.TLabel").pack(side="left")
        self.count_label = ttk.Label(top, text="0 images", style="CardMuted.TLabel")
        self.count_label.pack(side="right")
        self.tabs = ttk.Notebook(card)
        self.tabs.grid(row=1, column=0, sticky="nsew")
        preview = ttk.Frame(self.tabs, padding=12, style="Card.TFrame")
        texttab = ttk.Frame(self.tabs, padding=12, style="Card.TFrame")
        self.tabs.add(preview, text="Image preview")
        self.tabs.add(texttab, text="Extracted text")
        preview.rowconfigure(0, weight=1)
        preview.columnconfigure(0, weight=1)
        self.preview_label = ttk.Label(preview, text="Choose an image to preview", anchor="center", style="CardMuted.TLabel")
        self.preview_label.grid(row=0, column=0, sticky="nsew")
        self.file_list.bind("<<ListboxSelect>>", self.show_preview)
        texttab.rowconfigure(0, weight=1)
        texttab.columnconfigure(0, weight=1)
        self.text_box = tk.Text(texttab, wrap="word", font=("Segoe UI", 10), relief="flat", bg="#0D1424", fg=self.INK, padx=12, pady=12, undo=True)
        self.text_box.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(texttab, orient="vertical", command=self.text_box.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.text_box.configure(yscrollcommand=scroll.set)
        controls = ttk.Frame(card, style="Card.TFrame")
        controls.grid(row=2, column=0, sticky="ew", pady=(13, 0))
        self.run_btn = ttk.Button(controls, text="Extract text", style="Accent.TButton", command=self.run_ocr)
        self.run_btn.pack(side="left")
        ttk.Button(controls, text="Copy text", style="Soft.TButton", command=self.copy_text).pack(side="left", padx=(8, 0))
        ttk.Button(controls, text="Export…", style="Soft.TButton", command=self.export).pack(side="left", padx=(8, 0))

    def _set_tesseract_status(self):
        detected = find_tesseract()
        if detected:
            pytesseract.pytesseract.tesseract_cmd = detected
            self.status_tess.configure(text="●  Tesseract ready", foreground=self.GREEN)
            self.engine_badge.configure(text="●  OCR ENGINE CONNECTED", style="Ready.TLabel")
        else:
            self.status_tess.configure(text="●  Tesseract not found", foreground="#D97706")
            self.engine_badge.configure(text="●  ENGINE NEEDS SETUP", style="Missing.TLabel")

    def choose_tesseract(self):
        path = filedialog.askopenfilename(title="Select tesseract.exe", filetypes=[("Tesseract executable", "tesseract.exe"), ("All files", "*.*")])
        if path:
            pytesseract.pytesseract.tesseract_cmd = path
            os.environ["TESSERACT_CMD"] = path
            self.status_tess.configure(text="●  Tesseract ready", foreground=self.GREEN)
            self.engine_badge.configure(text="●  OCR ENGINE CONNECTED", style="Ready.TLabel")

    def add_images(self):
        paths = filedialog.askopenfilenames(title="Choose images", filetypes=[("Image files", "*.png *.jpg *.jpeg *.tif *.tiff *.bmp *.gif *.webp"), ("All files", "*.*")])
        self._add_paths(paths)

    def add_folder(self):
        folder = filedialog.askdirectory(title="Choose image folder")
        if folder:
            self._add_paths(sorted(str(p) for p in Path(folder).iterdir() if p.is_file() and p.suffix.lower() in SUPPORTED))

    def _add_paths(self, paths):
        existing = set(self.paths)
        for path in paths:
            if path not in existing:
                self.paths.append(path)
                self.file_list.insert("end", Path(path).name)
                existing.add(path)
        self.count_label.configure(text=f"{len(self.paths)} image" + ("s" if len(self.paths) != 1 else ""))
        if self.paths and not self.file_list.curselection():
            self.file_list.selection_set(0)
            self.show_preview()

    def remove_selected(self):
        selected = list(self.file_list.curselection())
        for index in reversed(selected):
            self.file_list.delete(index)
            del self.paths[index]
        self.count_label.configure(text=f"{len(self.paths)} images")
        self.preview_label.configure(image="", text="Choose an image to preview")
        self.photo = None

    def show_preview(self, _event=None):
        selection = self.file_list.curselection()
        if not selection:
            return
        try:
            with Image.open(self.paths[selection[0]]) as img:
                img = ImageOps.exif_transpose(img).convert("RGB")
                img.thumbnail((690, 470))
                self.photo = ImageTk.PhotoImage(img)
                self.preview_label.configure(image=self.photo, text="")
        except Exception as e:
            self.preview_label.configure(image="", text=f"Preview unavailable: {e}")

    def _prepare(self, image, use_preprocessing):
        image = ImageOps.exif_transpose(image).convert("RGB")
        if use_preprocessing:
            image = ImageOps.grayscale(image)
            image = ImageOps.autocontrast(image)
            image = ImageEnhance.Contrast(image).enhance(1.25)
            image = image.filter(ImageFilter.MedianFilter(size=3))
        return image

    def run_ocr(self):
        if not self.paths:
            messagebox.showwarning("No images", "Add at least one image to get started.")
            return
        if self.busy:
            return
        self.busy = True
        self.results = []
        self.text_box.delete("1.0", "end")
        self.run_btn.configure(state="disabled")
        self.progress.configure(value=0)
        self.status.configure(text="Starting OCR…")
        threading.Thread(target=self._worker, args=(list(self.paths), self.lang.get().strip() or "eng", self._psm_value(), self.preprocess.get()), daemon=True).start()

    def _psm_value(self):
        return {"Auto page (3)": 3, "Single block (6)": 6, "Single line (7)": 7, "Sparse text (11)": 11}.get(self.psm.get(), 3)

    def _worker(self, paths, language, psm, use_preprocessing):
        results = []
        for i, path in enumerate(paths, 1):
            try:
                with Image.open(path) as source:
                    original = self._prepare(source, use_preprocessing)
                    variants = [original]
                    if not use_preprocessing:
                        enhanced = ImageOps.grayscale(original)
                        enhanced = ImageOps.autocontrast(enhanced)
                        enhanced = ImageEnhance.Contrast(enhanced).enhance(1.2)
                        enhanced = enhanced.filter(ImageFilter.SHARPEN)
                        longest = max(enhanced.size)
                        if longest < 1500:
                            scale = min(2.0, 1500 / max(1, longest))
                            enhanced = enhanced.resize((int(enhanced.width * scale), int(enhanced.height * scale)), Image.Resampling.LANCZOS)
                        variants.append(enhanced)

                    layouts = [psm]
                    if psm in (3, 6):
                        layouts = [psm, 6 if psm == 3 else 3]
                    best = None
                    for variant in variants:
                        for layout in layouts:
                            config = f"--oem 3 --psm {layout}"
                            text = pytesseract.image_to_string(variant, lang=language, config=config).strip()
                            data = pytesseract.image_to_data(variant, lang=language, config=config, output_type=Output.DICT)
                            confidences = [float(value) for value in data["conf"] if float(value) >= 0]
                            confidence = round(sum(confidences) / len(confidences)) if confidences else 0
                            words = [word for word in text.split() if any(char.isalnum() for char in word)]
                            score = confidence + min(len(words), 40) * 0.08 + min(len(text), 1000) / 1000
                            candidate = {"file": Path(path).name, "path": path, "text": text, "confidence": confidence, "error": "", "score": score}
                            if best is None or candidate["score"] > best["score"]:
                                best = candidate
                    if best is None:
                        best = {"file": Path(path).name, "path": path, "text": "", "confidence": 0, "error": "No OCR result was produced."}
                    best.pop("score", None)
                    results.append(best)
            except Exception as e:
                results.append({"file": Path(path).name, "path": path, "text": "", "confidence": 0, "error": str(e)})
            self.root.after(0, self._progress, i, len(paths), Path(path).name)
        self.root.after(0, self._finish, results)

    def _progress(self, done, total, name):
        self.progress.configure(value=done / total * 100)
        self.status.configure(text=f"Reading {name}  ·  {done} of {total}")

    def _finish(self, results):
        self.results = results
        self.busy = False
        self.run_btn.configure(state="normal")
        self.status.configure(text=f"Finished · {len(results)} image(s)")
        self.text_box.delete("1.0", "end")
        for item in results:
            self.text_box.insert("end", f"{item['file']}  ·  confidence {item['confidence']}%\n")
            if item["error"]:
                self.text_box.insert("end", f"ERROR: {item['error']}\n")
            else:
                self.text_box.insert("end", (item["text"] or "[No text detected]") + "\n")
            self.text_box.insert("end", "\n" + "─" * 50 + "\n\n")
        self.tabs.select(1)
        errors = [r for r in results if r["error"]]
        if errors and len(errors) == len(results):
            messagebox.showerror("OCR error", "Could not process selected image(s). Check OCR settings and Tesseract installation.\n\n" + errors[0]["error"])

    def copy_text(self):
        text = self.text_box.get("1.0", "end").strip()
        if not text:
            messagebox.showinfo("Nothing to copy", "Run OCR or type text in the preview first.")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.status.configure(text="Text copied to clipboard")

    def export(self):
        if not self.results:
            messagebox.showinfo("Nothing to export", "Run OCR first.")
            return
        path = filedialog.asksaveasfilename(title="Export OCR results", defaultextension=".txt", filetypes=[("Text file", "*.txt"), ("CSV spreadsheet", "*.csv"), ("Excel workbook", "*.xlsx")])
        if not path:
            return
        try:
            ext = Path(path).suffix.lower()
            if ext == ".csv":
                with open(path, "w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f)
                    writer.writerow(["File Name", "Confidence (%)", "Extracted Text", "Error"])
                    for r in self.results:
                        writer.writerow([r["file"], r["confidence"], r["text"], r["error"]])
            elif ext == ".xlsx":
                try:
                    import pandas as pd
                except ImportError:
                    messagebox.showerror("Excel support missing", "Install Excel export support with: pip install pandas openpyxl")
                    return
                pd.DataFrame([{"File Name": r["file"], "Confidence (%)": r["confidence"], "Extracted Text": r["text"], "Error": r["error"]} for r in self.results]).to_excel(path, index=False)
            else:
                with open(path, "w", encoding="utf-8") as f:
                    for r in self.results:
                        f.write(f"{r['file']}  |  confidence {r['confidence']}%\n{r['text']}\n\n")
            self.status.configure(text=f"Saved {Path(path).name}")
            messagebox.showinfo("Export complete", f"Results saved to:\n{path}")
        except Exception as e:
            messagebox.showerror("Export failed", str(e))


def main():
    root = tk.Tk()
    LumaOCR(root)
    root.mainloop()


if __name__ == "__main__":
    main()














