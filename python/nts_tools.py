import os
import fitz  # PyMuPDF
import subprocess
import tkinter as tk
from tkinter import filedialog
import re


def process_nts(
    wb,
    summary_ws,
    nts_folder,
    output_folder,
    merged_path,
    update_status,
    merge_pdfs,
    nts_pdf_folder,
):
    try:
        # === Read flags and sheet info ===
        target_sheet_name = summary_ws.Range("F5").Value
        target_ws = wb.Sheets(target_sheet_name)

        merge_flag = bool(summary_ws.Range("I26").Value)
        filter_order_pages = bool(summary_ws.Range("L26").Value)
        combine_flag = bool(summary_ws.Range("I27").Value)

        # === Collect NTS and Waybill numbers ===
        nts_list = []
        waybill_list = []
        row = 4
        while True:
            nts_no = target_ws.Cells(row, 2).Value
            wb_no = target_ws.Cells(row, 3).Value
            if not nts_no and not wb_no:
                break
            if nts_no:
                nts_list.append(str(nts_no).strip().replace(".0", ""))
            if wb_no:
                waybill_list.append(str(wb_no).strip().replace(".0", ""))
            row += 1

        if not nts_list:
            update_status("⚠️ No NTS numbers found")
            return False

        # === Validate merge + combine flags ===
        if not merge_flag and combine_flag:
            summary_ws.Range("I5").Value = "❌ Merge flag is OFF while combine is ON"
            update_status("Check (✔) NTS  [Merge + Open]  to continue..")
            return False

        # === Check all NTS files before doing anything ===
        missing_nts = []
        for i, nts_no in enumerate(nts_list, start=1):
            pdf_path = os.path.join(nts_folder, f"{nts_no}.pdf")
            if not os.path.exists(pdf_path):
                missing_nts.append(f"{i} - {nts_no}")

        if missing_nts:
            summary_ws.Range("T8:T100").ClearContents()
            summary_ws.Range("S5").Value = f"Query - {len(missing_nts)}"
            for r, miss in enumerate(missing_nts, start=8):
                summary_ws.Range(f"T{r}").Value = miss
            update_status("NTS PDF's not found..          ➤")
            return False

        # === Ask for waybill file if combining Waybill + NTS ===
        waybill_selected = False
        if combine_flag:
            root = tk.Tk()
            root.withdraw()
            waybill_file = filedialog.askopenfilename(
                title="Select Combined Waybill PDF (Cancel to skip)",
                filetypes=[("PDF files", "*.pdf")],
            )
            waybill_selected = bool(waybill_file)
            if not waybill_selected:
                summary_ws.Range("I5").Value = "WAYBILL NOT MATCHING"
                update_status("❌ No Waybill file selected")
                return False

        # === Helper: Extract waybill page ===
        def extract_waybill_page(waybill_number):
            doc = fitz.open(waybill_file)
            extracted = fitz.open()
            wb_clean = waybill_number.strip().replace(" ", "").replace("-", "")
            found_any = False

            for i, page in enumerate(doc):
                text = page.get_text()
                candidates = re.findall(r"\b[457]\d{9,14}\b", text)
                for candidate in candidates:
                    if candidate.strip() == wb_clean:
                        extracted.insert_pdf(doc, from_page=i, to_page=i)
                        found_any = True
                        break

            doc.close()
            return extracted if found_any else None


        # === Helper: Extract NTS delivery page ===
        def extract_nts_slip(nts_number, filter_pages):
            filename = nts_number + ".pdf"
            path = os.path.join(nts_folder, filename)
            if not os.path.exists(path):
                return None
            doc = fitz.open(path)
            extracted = fitz.open()

            if filter_pages:
                pattern = re.compile(
                    r"(order\s*slip\s*/\s*delivery\s*challan|delivery\s*challan\s*/\s*pick\s*list)",
                    re.IGNORECASE,
                )
                for i, page in enumerate(doc):
                    text = page.get_text()
                    if pattern.search(text):
                        extracted.insert_pdf(doc, from_page=i, to_page=i)
            else:
                extracted.insert_pdf(doc)

            doc.close()
            return extracted if len(extracted) > 0 else None

        # === Waybill + NTS Merge ===
        if combine_flag:
            final_doc = fitz.open()
            total_pairs = min(len(nts_list), len(waybill_list))

            for i in range(total_pairs):
                wb_num = waybill_list[i]
                nts_num = nts_list[i]

                waybill_pdf = extract_waybill_page(wb_num)
                if not waybill_pdf:
                    summary_ws.Range("I5").Value = f"Waybill not found: {wb_num}"
                    update_status("Waybill not found.. ➤")
                    return False

                nts_pdf = extract_nts_slip(nts_num, filter_order_pages)
                if not nts_pdf:
                    summary_ws.Range("I5").Value = f"NTS not found: {nts_num}"
                    update_status("NTS not found.. ➤")
                    return False

                final_doc.insert_pdf(waybill_pdf)
                final_doc.insert_pdf(nts_pdf)
                waybill_pdf.close()
                nts_pdf.close()

            final_doc.save(merged_path)
            final_doc.close()
            if filter_order_pages:
                update_status("🚀 Combined Waybills + Filtered NTS (Delivery Challan pages only)")
            else:
                update_status("🚀 Combined Waybills + Full NTS (All pages)")
            summary_ws.Range("I5").Value = "NTS Combine Completed.. 🚀"
            subprocess.Popen(
                [r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", merged_path]
            )
            return True

        # === NTS Only Merge or Export ===
        else:
            if merge_flag:
                merged_doc = fitz.open()
                for nts_no in nts_list:
                    pdf_path = os.path.join(nts_folder, f"{nts_no}.pdf")
                    doc = fitz.open(pdf_path)

                    if filter_order_pages:
                        for j in range(len(doc)):
                            text = doc[j].get_text().lower()
                            if "order slip" in text or "delivery challan" in text:
                                merged_doc.insert_pdf(doc, from_page=j, to_page=j)
                    else:
                        merged_doc.insert_pdf(doc)

                    doc.close()

                merged_doc.save(merged_path)
                merged_doc.close()
                subprocess.Popen(
                    [r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", merged_path]
                )
                summary_ws.Range("I5").Value = "NTS Combine Completed.. 🚀"
                return True

            else:
                # Export individually
                for i, nts_no in enumerate(nts_list, start=1):
                    pdf_path = os.path.join(nts_folder, f"{nts_no}.pdf")
                    doc = fitz.open(pdf_path)
                    filtered_doc = fitz.open()

                    if filter_order_pages:
                        for j in range(len(doc)):
                            text = doc[j].get_text().lower()
                            if "order slip" in text or "delivery challan" in text:
                                filtered_doc.insert_pdf(doc, from_page=j, to_page=j)
                    else:
                        filtered_doc.insert_pdf(doc)

                    dest = os.path.join(output_folder, f"{i}_{nts_no}.pdf")
                    filtered_doc.save(dest)
                    doc.close()
                    filtered_doc.close()

                summary_ws.Range("I5").Value = "NTS Export Completed.. 📂"
                return True

    except Exception as e:
        update_status(f"❌ Error in process_nts: {e}")
        summary_ws.Range("I5").Value = f"ERROR: {e}"
        return False
