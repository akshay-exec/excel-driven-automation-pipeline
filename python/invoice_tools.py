import os
import fitz  # PyMuPDF
import subprocess


def process_invoices(wb, summary_ws, invoice_folder, output_folder, merged_path, update_status, merge_pdfs):
    try:
        # --- Get target sheet ---
        target_sheet_name = summary_ws.Range("F5").Value
        target_ws = wb.Sheets(target_sheet_name)

        # --- Flags from Set sheet ---
        merge_flag = bool(summary_ws.Range("I22").Value)   # Merge PDFs?
        remove_barcode_pages = bool(summary_ws.Range("L22").Value)  # Remove pages containing "barcode"

        # --- Collect serial numbers & invoice numbers ---
        invoices = []  # [(serial, inv_no, cdo_no)]
        row = 4
        while True:
            serial = target_ws.Range(f"A{row}").Value
            inv_no = target_ws.Range(f"B{row}").Value
            cdo_no = target_ws.Range(f"I{row}").Value

            if not inv_no:
                break

            if serial is not None:
                try:
                    serial = str(int(float(serial)))
                except Exception:
                    serial = str(serial).strip()
            else:
                serial = str(row - 3)

            invoices.append(
                (serial, str(inv_no).strip(), str(cdo_no).strip() if cdo_no else None)
            )
            row += 1

        if not invoices:
            update_status("⚠️ No invoice numbers found in target sheet")
            return False

        # --- Check for missing PDFs ---
        pdfs = []
        missing_invoices = []

        for serial, inv_no, cdo_no in invoices:
            pdf_path = os.path.join(invoice_folder, f"{inv_no}.pdf")
            if os.path.exists(pdf_path):
                pdfs.append((serial, pdf_path))
            else:
                if cdo_no:
                    missing_invoices.append(f"{serial} - {inv_no} CDO ({cdo_no})")
                else:
                    missing_invoices.append(f"{serial} - {inv_no}")

        if missing_invoices:
            summary_ws.Range("S5").Value = f"Query - {len(missing_invoices)}"
            r_row = 8
            for miss in missing_invoices:
                summary_ws.Range(f"T{r_row}").Value = miss
                r_row += 1
            update_status("E-invoice not found..          ➤")
            return False

        # --- Merge PDFs if merge_flag is True ---
        if merge_flag:
            merged_doc = fitz.open()
            for serial, pdf in pdfs:
                doc = fitz.open(pdf)
                pages_to_keep = []

                if remove_barcode_pages:
                    for i in range(len(doc)):
                        text = doc[i].get_text().lower()
                        if "barcode" not in text:
                            pages_to_keep.append(i)
                else:
                    pages_to_keep = list(range(len(doc)))

                for p in pages_to_keep:
                    merged_doc.insert_pdf(doc, from_page=p, to_page=p)
                doc.close()

            merged_doc.save(merged_path)
            merged_doc.close()
            subprocess.Popen(
                [r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", merged_path]
            )

        else:  # Export individually
            for i, (serial, pdf) in enumerate(pdfs, start=1):
                doc = fitz.open(pdf)
                pages_to_keep = []

                if remove_barcode_pages:
                    for j in range(len(doc)):
                        text = doc[j].get_text().lower()
                        if "barcode" not in text:
                            pages_to_keep.append(j)
                else:
                    pages_to_keep = list(range(len(doc)))

                filtered_doc = fitz.open()
                for p in pages_to_keep:
                    filtered_doc.insert_pdf(doc, from_page=p, to_page=p)

                dest_name = f"{i}_{os.path.splitext(os.path.basename(pdf))[0]}.pdf"
                dest_path = os.path.join(output_folder, dest_name)
                filtered_doc.save(dest_path)
                filtered_doc.close()
                doc.close()

            update_status(
                "Export done: Invoices copied"
                + (" (barcode pages removed)" if remove_barcode_pages else "")
            )

        return True

    except Exception as e:
        update_status(f"❌ Error in process_invoices: {e}")
        return False
