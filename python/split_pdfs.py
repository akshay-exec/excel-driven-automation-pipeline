import os
import re
import pdfplumber
from PyPDF2 import PdfReader, PdfWriter
from collections import defaultdict

# Waybill number pattern: starts with 4, 5, or 7 and is 10–15 digits long
WAYBILL_PATTERN = r'\b[457]\d{9,14}\b'

def extract_waybill_from_pdf_page(pdf_path, page_number):
    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[page_number].extract_text()
        if text:
            match = re.search(WAYBILL_PATTERN, text)
            if match:
                return match.group(0)
    return None

def split_group_and_move(folder_path, output_folder):
    # Creates output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)

    for filename in os.listdir(folder_path):
        if filename.lower().endswith(".pdf"):
            file_path = os.path.join(folder_path, filename)
            try:
                pdf_reader = PdfReader(file_path)
                num_pages = len(pdf_reader.pages)

                # Dictionary to collect pages by waybill
                waybill_pages = defaultdict(list)

                for page_num in range(num_pages):
                    waybill = extract_waybill_from_pdf_page(file_path, page_num)
                    if waybill:
                        waybill_pages[waybill].append(pdf_reader.pages[page_num])
                    else:
                        # Fallback: Save unidentified page
                        fallback_writer = PdfWriter()
                        fallback_writer.add_page(pdf_reader.pages[page_num])
                        fallback_filename = f"{os.path.splitext(filename)[0]}_page_{page_num+1}_unidentified.pdf"
                        fallback_path = os.path.join(output_folder, fallback_filename)
                        with open(fallback_path, 'wb') as f:
                            fallback_writer.write(f)
                        print(f"Saved unidentified page as: {fallback_filename}")

                # Save grouped pages by waybill
                for waybill, pages in waybill_pages.items():
                    writer = PdfWriter()
                    for p in pages:
                        writer.add_page(p)

                    output_path = os.path.join(output_folder, f"{waybill}.pdf")
                    with open(output_path, 'wb') as f:
                        writer.write(f)
                    print(f"Saved: {waybill}.pdf with {len(pages)} page(s)")

                # ✅ Delete the original master file
                os.remove(file_path)
                # print(f"Deleted original file: {filename}")

            except Exception as e:
                print(f"Error processing {filename}: {e}")

# 🔧 Update these paths inside the script before running
input_folder = r"PATH_TO_INPUT_PDF_FOLDER"
output_folder = r"PATH_TO_OUTPUT_PDF_FOLDER"

split_group_and_move(input_folder, output_folder)
