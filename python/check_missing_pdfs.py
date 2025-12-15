import os
import sys
import pythoncom
import win32com.client as win32

#-----Update these paths inside the script before running-----
EXCEL_PATH = r"PATH_TO_EXCEL_FILE.xlsm"
SHEET_NAME = "SHEET_NAME"
PDF_FOLDER_PATH = r"PATH_TO_PDF_FOLDER"

def clean_value(val):
    """Convert cell value to clean string without float suffixes like '.0'."""
    if val is None:
        return None
    val = str(val).strip()
    if val.endswith(".0"):
        val = val[:-2]
    return val

def check_missing_pdfs():
    pythoncom.CoInitialize()
    try:
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        wb = excel.Workbooks.Open(EXCEL_PATH)
        ws = wb.Sheets(SHEET_NAME)

        # Read values from columns A and B
        row = 2
        invoice_numbers = []
        waybill_numbers = []

        while True:
            invoice = clean_value(ws.Cells(row, 1).Value)  # Column A
            waybill = clean_value(ws.Cells(row, 2).Value)  # Column B

            if not invoice and not waybill:
                break
            if invoice:
                invoice_numbers.append(invoice)
            if waybill:
                waybill_numbers.append(waybill)
            row += 1

        all_numbers = invoice_numbers + waybill_numbers
        total = len(all_numbers)
        missing_files = []

        bar_length = 40
        print(f"Checking {total} file(s)...")

        for i, number in enumerate(all_numbers, start=1):
            expected_file = os.path.join(PDF_FOLDER_PATH, f"{number}.pdf")

            if not os.path.isfile(expected_file):
                missing_files.append(number)

            # Progress bar
            progress_percent = int((i / total) * 100)
            filled_length = int(bar_length * progress_percent // 100)
            bar = '=' * filled_length + '-' * (bar_length - filled_length)
            print(f"\rChecking files: |{bar}| {progress_percent}%", end="")
            sys.stdout.flush()

            # Update progress in Excel cell J3
            ws.Range("J3").Value = f"▷ {progress_percent}%"
            #wb.Save()

        # Output missing filenames in T5 downward
        start_row = 5
        if missing_files:
            for idx, filename in enumerate(missing_files):
                ws.Cells(start_row + idx, 20).Value = filename  # Column T
            ws.Range("S3").Value = "Missing Files"
        else:
            ws.Range("S3").Value = "All PDF found.."

        print("\n\nCheck complete.")
        if missing_files:
            print(f"Missing {len(missing_files)} PDF file(s):")
            for f in missing_files:
                print(f)
        else:
            print("All PDF files are present.")

        wb.Save()

    except Exception as e:
        print(f"Error: {e}")
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    check_missing_pdfs()
