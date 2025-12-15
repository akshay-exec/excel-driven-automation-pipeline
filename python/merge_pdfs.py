import os
from PyPDF2 import PdfMerger
import tkinter as tk
from tkinter import filedialog

# === Update these paths inside the script before running ===
pdf_folder = r"PATH_TO_PDF_FOLDER"

# === Initialize Tkinter root (hide main window) ===
root = tk.Tk()
root.withdraw()  # Hide the root window

# === Open Save As dialog for user to pick location and filename ===
output_file = filedialog.asksaveasfilename(
    defaultextension=".pdf",
    filetypes=[("PDF files", "*.pdf")],
    initialdir=pdf_folder,
    title="Save merged PDF as"
)

if not output_file:
    print("No file selected, exiting...")
    exit()

# === Get PDF files in order of numeric prefix ===
pdf_files = []
for fname in os.listdir(pdf_folder):
    if fname.lower().endswith(".pdf"):
        try:
            number = int(fname.split('_')[0])
            pdf_files.append((number, fname))
        except ValueError:
            pass

pdf_files.sort(key=lambda x: x[0])

# === Merge PDFs ===
merger = PdfMerger()
for _, fname in pdf_files:
    merger.append(os.path.join(pdf_folder, fname))

merger.write(output_file)
merger.close()

print(f"Merged PDF saved to: {output_file}")
