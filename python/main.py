import os
import sys
try:
    import win32com.client
except ImportError:
    print("win32com.client left the chat! ☹️")
    sys.exit(1)
import shutil
from modules import invoice_tools as it
from modules import nts_tools as nt

# === PATHS ===
excel_path = r"EXCEL.xlsm"

invoice_folder = r"FOLDER_PATH"
nts_folder = r"FOLDER_PATH"

output_folder_inv = r"FOLDER_PATH"
output_folder_nts = r"FOLDER_PATH"

merged_path_inv = os.path.join(output_folder_inv, "Merged_Invoices.pdf")
merged_path_nts = os.path.join(output_folder_nts, "Merged_NTS.pdf")

os.makedirs(output_folder_inv, exist_ok=True)
os.makedirs(output_folder_nts, exist_ok=True)

# === Utility: clean a folder ===
def clean_folder(folder):
    for f in os.listdir(folder):
        path = os.path.join(folder, f)
        try:
            if os.path.isfile(path):
                os.remove(path)
            elif os.path.isdir(path):
                shutil.rmtree(path)
        except Exception as e:
            print(f"⚠️ Could not remove {path}: {e}")


# === Excel connection ===
excel = win32com.client.Dispatch("Excel.Application")
wb = excel.Workbooks.Open(excel_path)
summary_ws = wb.Sheets("Set")

def update_status(msg: str):
    """Update Excel + console"""
    summary_ws.Range("I5").Value = msg
    print(msg)


# === Always clean both output folders before starting ===
clean_folder(output_folder_inv)
clean_folder(output_folder_nts)


# === Mode Selection ===
mode = sys.argv[1] if len(sys.argv) > 1 else None

if mode == "INV":
    flag = bool(summary_ws.Range("I22").Value)
    update_status("Processing INV... ⏳")

    success = it.process_invoices(
        wb=wb,
        summary_ws=summary_ws,
        invoice_folder=invoice_folder,
        output_folder=output_folder_inv,
        merged_path=merged_path_inv,
        update_status=update_status,
        merge_pdfs=flag
    )

    if success:
        if flag:
            update_status("Invoice Combine Completed.. 🚀")   # when merged
        else:
            update_status("Export Completed.. ✅")       # when exported only


elif mode == "NTS":
    flag = bool(summary_ws.Range("I26").Value)
    update_status("Processing NTS... ⏳")

    success = nt.process_nts(
        wb=wb,
        summary_ws=summary_ws,
        nts_folder=nts_folder,
        output_folder=output_folder_nts,
        merged_path=merged_path_nts,
        update_status=update_status,
        merge_pdfs=flag,
        nts_pdf_folder=nts_folder
    )

    if success:
        if flag:
            update_status("NTS Combine Completed.. 🚀")   # when merged
        else:
            update_status("Export Completed.. ✅")       # when exported only


else:
    update_status("❌ No mode specified (use INV or NTS)")
