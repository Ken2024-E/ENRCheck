# noinspection PyInterpreter
import os
import re
import sys
import platform
import threading
from collections import Counter
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from docx2pdf import convert
import fitz  # PyMuPDF
from PIL import Image
import io

# --- Robust Date Parsing Module ---
try:
    from dateutil import parser as date_parser
except ImportError:
    date_parser = None

# --- OCR System Mapping Integration ---
try:
    import pytesseract

    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False


class DocVerifierApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Document Integrity & Body Structure Verifier")
        self.root.geometry("750x600")
        self.root.minsize(650, 500)

        # Style configurations
        self.style = ttk.Style()
        self.style.theme_use("clam")

        self.create_widgets()

    def create_widgets(self):
        # --- File Selection Frame ---
        file_frame = ttk.LabelFrame(self.root, text=" Document Selection ", padding=15)
        file_frame.pack(fill="x", padx=15, pady=10)

        ttk.Label(file_frame, text="Source Word Document (.docx):").grid(
            row=0, column=0, sticky="w", pady=2
        )

        self.file_path_var = tk.StringVar()
        self.file_entry = ttk.Entry(
            file_frame, textvariable=self.file_path_var, width=55
        )
        self.file_entry.grid(row=1, column=0, padx=(0, 10), sticky="ew")

        self.browse_btn = ttk.Button(
            file_frame, text="Browse...", command=self.browse_file
        )
        self.browse_btn.grid(row=1, column=1, sticky="e")

        file_frame.columnconfigure(0, weight=1)

        # --- Control & Progress Frame ---
        control_frame = ttk.Frame(self.root, padding=5)
        control_frame.pack(fill="x", padx=15, pady=5)

        self.run_btn = ttk.Button(
            control_frame, text="Analyze Document Layout", command=self.start_analysis_thread
        )
        self.run_btn.pack(side="left", padx=(0, 15))

        self.progress_bar = ttk.Progressbar(control_frame, mode="indeterminate", length=150)
        self.status_label = ttk.Label(control_frame, text="Ready", font=("Arial", 9, "italic"))
        self.status_label.pack(side="left", fill="x", expand=True)

        # --- Terminal/Log Output Frame ---
        log_frame = ttk.LabelFrame(self.root, text=" Document Analysis Logs ", padding=10)
        log_frame.pack(fill="both", expand=True, padx=15, pady=(5, 15))

        self.log_text = tk.Text(
            log_frame, wrap="word", background="#2b2b2b", foreground="#f0f0f0", font=("Consolas", 10)
        )
        self.log_text.pack(side="left", fill="both", expand=True)

        # Add scrollbar to text log
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.log_text.config(yscrollcommand=scrollbar.set)

        if not OCR_AVAILABLE:
            self.log_print(
                "⚠ SYSTEM WARNING: 'pytesseract' module was not detected.\n"
                "  To read headings embedded as images/screenshots, run:\n"
                "  'pip install pytesseract pillow' and confirm Tesseract installation."
            )

        if date_parser is None:
            self.log_print(
                "⚠ DEPENDENCY WARNING: 'python-dateutil' module not detected.\n"
                "  For reliable multi-format date verification, please run:\n"
                "  'pip install python-dateutil'"
            )

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select Word Document", filetypes=[("Word Documents", "*.docx")]
        )
        if filename:
            self.file_path_var.set(filename)
            self.log_clear()
            self.log_print(f"Loaded: {filename}\nReady for structural verification.")

    def log_print(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)

    def log_clear(self):
        self.log_text.delete("1.0", tk.END)

    def update_status(self, text, start_progress=False, stop_progress=False):
        self.status_label.config(text=text)
        if start_progress:
            self.progress_bar.pack(side="right", padx=5)
            self.progress_bar.start(10)
        if stop_progress:
            self.progress_bar.stop()
            self.progress_bar.pack_forget()

    def start_analysis_thread(self):
        docx_path = self.file_path_var.get()
        if not docx_path or not os.path.exists(docx_path):
            messagebox.showerror(
                "Error", "Please select a valid .docx file before running analysis."
            )
            return

        self.run_btn.config(state="disabled")
        self.browse_btn.config(state="disabled")
        self.file_entry.config(state="disabled")
        self.update_status("Processing... Extracting Document Structures", start_progress=True)
        self.log_clear()

        base_dir = os.path.dirname(docx_path)
        base_name = os.path.splitext(os.path.basename(docx_path))[0]
        final_pdf_path = os.path.join(base_dir, f"{base_name}_layout_verified.pdf")

        threading.Thread(
            target=self.run_core_logic,
            args=(docx_path, final_pdf_path),
            daemon=True,
        ).start()

    def check_header_revision_consistency(self, doc, body_enr_baseline):
        """Isolates the top 12% height and right 33% width of the page header zone.
        Captures the standalone single-digit revision from the 3rd page header
        to use as the official Baseline ENR Revision Number, then checks all other
        pages for any inconsistent standalone values, highlighting mismatches in yellow."""
        self.log_print("\n[Executing Dedicated Header Revision Verification Scan]")

        baseline_rev_val = None
        single_digit_pattern = re.compile(r'^\s*([A-Z0-9])\s*$', re.IGNORECASE)

        # --- STEP 1: CAPTURE THE BASELINE FROM THE 3RD PAGE HEADER ---
        if len(doc) >= 3:
            page_3 = doc[2]  # Index 2 represents the 3rd page
            p3_width = page_3.rect.width
            p3_right_side = p3_width * (1.0 - 0.33)
            p3_header_zone = fitz.Rect(p3_right_side, 0, p3_width, page_3.rect.height * 0.12)

            p3_text = page_3.get_text("text", clip=p3_header_zone).strip()

            # OCR Fallback for Page 3 header if layout text layer is empty
            if not p3_text and OCR_AVAILABLE:
                try:
                    pix = page_3.get_pixmap(matrix=fitz.Matrix(3, 3), clip=p3_header_zone)
                    img = Image.open(io.BytesIO(pix.tobytes("png")))
                    p3_text = pytesseract.image_to_string(img).strip()
                except Exception:
                    pass

            # Parse lines strictly looking for the standalone single-digit revision number
            for line in p3_text.splitlines():
                match = single_digit_pattern.match(line.strip())
                if match:
                    baseline_rev_val = match.group(1).strip().upper()
                    self.log_print(
                        f"-> Captured Baseline ENR Revision Number from Page 3 Header: Rev {baseline_rev_val}")
                    break

        # --- STEP 2: SAFETY FALLBACK IF 3RD PAGE PARSING FAILS ---
        if not baseline_rev_val:
            self.log_print("-> Warning: Could not isolate a clean, single-digit revision value from Page 3 Header.")
            if body_enr_baseline:
                body_rev_match = re.search(r'(?:Rev|Revision)\s*\.?\s*([A-Z0-9]+)', body_enr_baseline, re.IGNORECASE)
                baseline_rev_val = body_rev_match.group(
                    1).strip().upper() if body_rev_match else body_enr_baseline.upper()
                self.log_print(
                    f"-> Fallback applied: Using baseline parsed from document body text: Rev {baseline_rev_val}")
            else:
                self.log_print(
                    "-> Skipped: No valid baseline value found on Page 3 or within the body text baseline layer.")
                return

        self.log_print(
            f"   Document baseline is set to version '{baseline_rev_val}'. Scanning other page headers for discrepancies...")
        deviations_found = False

        # General pattern to find standalone single alphanumeric characters in the headers
        header_value_pattern = re.compile(r'\b[A-Z0-9]\b', re.IGNORECASE)

        # --- STEP 3: EVALUATE ALL OTHER PAGE HEADERS AGAINST THE BASELINE ---
        for page_num in range(len(doc)):
            # Skip the control page (Page 3) to prevent it from flagging itself
            if page_num == 2:
                continue

            page = doc[page_num]
            right_side_start = page.rect.width * (1.0 - 0.33)
            header_zone = fitz.Rect(right_side_start, 0, page.rect.width, page.rect.height * 0.12)
            header_text = page.get_text("text", clip=header_zone)

            if not header_text.strip() and OCR_AVAILABLE:
                try:
                    pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), clip=header_zone)
                    img = Image.open(io.BytesIO(pix.tobytes("png")))
                    header_text = pytesseract.image_to_string(img)
                except Exception:
                    pass

            # Search text layout layers for standalone single character blocks
            for match in header_value_pattern.finditer(header_text):
                found_val = match.group(0).strip().upper()

                # Check all bounding matches inside this page's header segment
                rects = page.search_for(found_val, clip=header_zone)
                for rect in rects:
                    valid_rect = rect & page.rect
                    if not valid_rect.is_empty:
                        try:
                            box_text = page.get_text("text", clip=valid_rect).strip().upper()

                            # If it's a strict single-character token and doesn't match our Page 3 baseline
                            if box_text == found_val and box_text != baseline_rev_val:
                                deviations_found = True
                                self.log_print(
                                    f"   ❌ Revision Mismatch on Page {page_num + 1}! Found '{box_text}' (Expected Baseline '{baseline_rev_val}')")

                                annot = page.add_highlight_annot(valid_rect)
                                if annot:
                                    annot.set_colors(stroke=(1, 1, 0))  # Apply yellow highlight to the variation
                                    annot.update()
                        except Exception:
                            pass

        if not deviations_found:
            self.log_print("   ✓ Success! All evaluated page headers are consistent with the Page 3 baseline.")

    def check_pattern_consistency(self, doc, regex_pattern, prefix_label):
        """Scans the doc for a given regex pattern, calculates baseline, and highlights mismatches."""
        self.log_print(f"\n[Checking '{prefix_label}' Prefix & Suffix Consistency]")

        all_instances = []
        for page_num, page in enumerate(doc):
            text = page.get_text("text")
            for match in regex_pattern.finditer(text):
                full_string = match.group(0).strip()
                normalized = re.sub(r'\s+', ' ', full_string).upper()
                all_instances.append((full_string, normalized, page_num))

        if not all_instances:
            self.log_print(f"-> No variants of '{prefix_label}' found.")
            return None

        just_normalized_strings = [item[1] for item in all_instances]
        counts = Counter(just_normalized_strings)
        canonical_code, canonical_count = counts.most_common(1)[0]
        self.log_print(f"-> Detected Baseline Code: {canonical_code} (Found {canonical_count} times)")

        mismatches = [code for code in counts if code != canonical_code]
        if not mismatches:
            self.log_print(f"   ✓ Success! All instances of {prefix_label} match perfectly.")
        else:
            self.log_print(f"   ⚠ Found inconsistent variations: {mismatches}")
            self.log_print("   Applying yellow highlights to deviations...")

            for raw_text, normalized_text, p_num in all_instances:
                if normalized_text != canonical_code:
                    page = doc[p_num]
                    rects = page.search_for(raw_text)
                    for rect in rects:
                        valid_rect = rect & page.rect
                        if not valid_rect.is_empty:
                            annot = page.add_highlight_annot(valid_rect)
                            if annot:
                                annot.set_colors(stroke=(1, 1, 0))
                                annot.update()
                    self.log_print(f"   - Highlighted mismatch '{raw_text.strip()}' on Page {p_num + 1}")

        return canonical_code

    def parse_date(self, date_str):
        if date_parser is None:
            for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%d-%b-%Y", "%d/%m/%Y"):
                try:
                    from datetime import datetime
                    return datetime.strptime(date_str.strip(), fmt)
                except ValueError:
                    continue
            return None
        try:
            return date_parser.parse(date_str, fuzzy=True)
        except (ValueError, OverflowError):
            return None

    def check_calibration_dates(self, doc):
        self.log_print("\n[Executing Structural Table-Based Calibration Validation]")
        date_pattern = re.compile(
            r"(\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2},?\s+\d{4}\b|"
            r"\b\d{1,4}[/\-\.]\d{1,2}[/\-\.]\d{2,4}\b)", re.IGNORECASE
        )
        tables_found_count = 0
        verified_rows_count = 0

        for page_num, page in enumerate(doc):
            tables = page.find_tables()
            if not tables:
                continue
            for table in tables:
                if table.col_count != 5:
                    continue
                matrix = table.extract()
                if not matrix or len(matrix) < 2:
                    continue
                header_row = [str(cell).strip().upper() for cell in matrix[0] if cell]
                matches_header = any("EQUIPMENT" in h for h in header_row) and \
                                 any("MAKE" in h or "MODEL" in h for h in header_row) and \
                                 any("SN/" in h or "ASSET" in h for h in header_row) and \
                                 any("CALIBRATION DATE" in h or ("CALIBRATION" in h and "DATE" in h) for h in
                                     header_row)
                if not matches_header:
                    continue

                tables_found_count += 1
                self.log_print(f"-> Target Equipment Matrix detected on Page {page_num + 1}. Validating data rows...")

                for row_idx, row_data in enumerate(matrix[1:], start=1):
                    cleaned_row = [str(cell).strip() for cell in row_data if cell]
                    row_string_dump = " ".join(cleaned_row)
                    if "N/A" in row_string_dump.upper() or not row_string_dump.strip():
                        continue
                    found_dates = date_pattern.findall(row_string_dump)
                    if len(found_dates) >= 2:
                        verified_rows_count += 1
                        cal_date_str, due_date_str = found_dates[0], found_dates[1]
                        cal_date, due_date = self.parse_date(cal_date_str), self.parse_date(due_date_str)
                        if cal_date and due_date:
                            try:
                                expected_year = cal_date.year + 1
                                if cal_date.month == 2 and cal_date.day == 29:
                                    is_match = (
                                                due_date.year == expected_year and due_date.month == 2 and due_date.day in (
                                        28, 29))
                                else:
                                    is_match = (
                                                due_date.year == expected_year and due_date.month == cal_date.month and due_date.day == cal_date.day)
                            except Exception:
                                is_match = False

                            if not is_match:
                                self.log_print(
                                    f"   ❌ Calibration Mismatch [Row {row_idx}]: Expected 1 Year separation, found: Cal={cal_date_str} | Due={due_date_str}")
                                try:
                                    row_bbox = table.rows[row_idx].bbox
                                    annot = page.add_highlight_annot(row_bbox)
                                    if annot:
                                        annot.set_colors(stroke=(1, 1, 0))
                                        annot.update()
                                except Exception:
                                    rects = page.search_for(cal_date_str) + page.search_for(due_date_str)
                                    for rect in rects:
                                        valid_rect = rect & page.rect
                                        if not valid_rect.is_empty:
                                            annot = page.add_highlight_annot(valid_rect)
                                            if annot:
                                                annot.set_colors(stroke=(1, 1, 0))
                                                annot.update()
        if tables_found_count == 0:
            self.log_print("❌ Search Warning: No 5-column tables matching designated header template were found.")
        else:
            self.log_print(
                f"✓ Table Analysis complete. Checked {tables_found_count} table(s) across {verified_rows_count} row metrics.")

    def extract_ocr_text_from_page(self, page):
        try:
            pix = page.get_pixmap(matrix=fitz.Matrix(4, 4))
            img_data = pix.tobytes("png")
            img = Image.open(io.BytesIO(img_data))
            return pytesseract.image_to_string(img)
        except Exception:
            return ""

    def open_pdf_automatically(self, file_path):
        try:
            self.log_print(f"-> Launching system viewer for {os.path.basename(file_path)}...")
            current_sys = platform.system()
            if current_sys == "Windows":
                os.startfile(file_path)
            elif current_sys == "Darwin":
                os.system(f'open "{file_path}"')
            else:
                os.system(f'xdg-open "{file_path}"')
        except Exception as e:
            self.log_print(f"⚠ Unable to launch default viewer automatically: {str(e)}")

    def run_core_logic(self, docx_path, final_pdf_path):
        temp_pdf = os.path.join(os.path.dirname(docx_path), "~temp_verification_layout.pdf")
        try:
            self.log_print("Step 1: Spawning Word engine to compute document layouts...")
            convert(docx_path, temp_pdf)
            self.log_print("✓ Conversion successful.")

            self.log_print("\nStep 2: Analyzing layout layers...")
            doc = fitz.open(temp_pdf)

            # --- FEATURE A: Prefix Consistency Scans ---
            enr_pattern = re.compile(r"\bENR-\d+(?:-[A-Z])?(?:\s+(?:Rev|Revision)\s*\.?\s*\d+)?", re.IGNORECASE)
            cs_pattern = re.compile(r"\bCS-\d+(?:-[A-Z])?(?:\s+(?:Rev|Revision)\s*\.?\s*\d+)?", re.IGNORECASE)

            body_enr_baseline = self.check_pattern_consistency(doc, enr_pattern, "ENR-xxxxx")
            self.check_pattern_consistency(doc, cs_pattern, "CS-xxxx")

            # --- EXPLICITLY PRINT OUT BASELINE ENR REVISION VALUE ---
            if body_enr_baseline:
                rev_extract = re.search(r'(REV(?:ISION)?\s*\.?\s*\d+)', body_enr_baseline, re.IGNORECASE)
                most_common_rev = rev_extract.group(1) if rev_extract else "No explicit suffix"
                self.log_print(f"\n📢 [SUMMARY] Most Common ENR Revision Document-Wide: {most_common_rev}")
            else:
                most_common_rev = "None Found"
                self.log_print(f"\n📢 [SUMMARY] Most Common ENR Revision Document-Wide: None Found")

            # --- EXECUTE HEADER REVISION SCAN ---
            self.check_header_revision_consistency(doc, body_enr_baseline)

            # --- FEATURE B: Table of Contents Listing Extraction ---
            self.log_print("\n[Extracting Clean Table of Contents Items]")
            toc_page_indices = []
            pages_to_check = min(15, len(doc))
            for page_num in range(pages_to_check):
                page_text = doc[page_num].get_text("text").upper()
                if "TABLE OF CONTENTS" in page_text or "REVISION HISTORY" in page_text:
                    toc_page_indices.append(page_num)
            if not toc_page_indices:
                toc_page_indices = [0, 1, 2]

            flexible_toc_line_pattern = re.compile(r"^(.*?)(?:[\s\.\-\_\t]{2,})?(\d+)\s*$", re.MULTILINE)
            raw_collected_items = []

            for idx in toc_page_indices:
                if idx >= len(doc):
                    continue
                page = doc[idx]
                page_text = page.get_text("text")
                if not page_text.strip() and OCR_AVAILABLE:
                    page_text = self.extract_ocr_text_from_page(page)

                lines = page_text.splitlines()
                for line in lines:
                    clean_line = line.strip()
                    if not clean_line or "TABLE OF CONTENTS" in clean_line.upper():
                        continue
                    if "ERROR! BOOKMARK NOT DEFINED" in clean_line.upper():
                        self.log_print(f"   ❌ TOC Broken Reference found on Page {idx + 1}: \"{clean_line}\"")
                        error_rects = page.search_for(line) or page.search_for("Error! Bookmark not defined.")
                        for rect in error_rects:
                            valid_rect = rect & page.rect
                            if not valid_rect.is_empty:
                                try:
                                    annot = page.add_highlight_annot(valid_rect)
                                    if annot:
                                        annot.set_colors(stroke=(1, 1, 0))
                                        annot.update()
                                except Exception:
                                    pass
                        continue

                    match = flexible_toc_line_pattern.match(clean_line)
                    if match:
                        raw_title = match.group(1).strip()
                        clean_title_no_dots = re.sub(r'[\s\.\-\_\t]+$', '', raw_title).strip()
                        if len(clean_title_no_dots) > 1 and not clean_title_no_dots.isdigit():
                            raw_collected_items.append(clean_title_no_dots)

            self.log_print(f"-> Total Clean Table of Contents rows identified: {len(raw_collected_items)}")
            sliced_search_targets = raw_collected_items[2:-2] if len(raw_collected_items) > 4 else []

            body_content_by_page = {}
            for page_num in range(len(doc)):
                if page_num in toc_page_indices:
                    continue
                p_text = doc[page_num].get_text("text")
                if not p_text.strip() and OCR_AVAILABLE:
                    p_text = self.extract_ocr_text_from_page(doc[page_num])
                body_content_by_page[page_num + 1] = p_text.lower()

            failed_search_detected = False
            for item in sliced_search_targets:
                item_lower = item.lower()
                match_found_globally = any(item_lower in content_dump for content_dump in body_content_by_page.values())
                if not match_found_globally:
                    failed_search_detected = True

            if failed_search_detected:
                for idx in toc_page_indices:
                    if idx >= len(doc):
                        continue
                    page = doc[idx]
                    heading_rects = page.search_for("Table of Contents") or page.search_for("TABLE OF CONTENTS")
                    for rect in heading_rects:
                        valid_rect = rect & page.rect
                        if not valid_rect.is_empty:
                            try:
                                annot = page.add_highlight_annot(valid_rect)
                                if annot:
                                    annot.set_colors(stroke=(1, 1, 0))
                                    annot.update()
                            except Exception:
                                pass

            # --- FEATURE C: Equipment Table Calibration Scans ---
            self.check_calibration_dates(doc)

            # --- FEATURE D: Heading Level Structural Parsing ---
            self.log_print("\n[Executing Body Structural Section Heading Verification]")
            heading_pattern = re.compile(r"^((?:\d+\.)+\d+)\s+(.+)$", re.MULTILINE)
            for page_num in range(len(doc)):
                if page_num in toc_page_indices:
                    continue
                page = doc[page_num]
                page_text = page.get_text("text")
                if (not page_text.strip() or "7.2." not in page_text) and OCR_AVAILABLE:
                    ocr_backup_text = self.extract_ocr_text_from_page(page)
                    if len(ocr_backup_text.strip()) > len(page_text.strip()):
                        page_text = ocr_backup_text
                for match in heading_pattern.finditer(page_text):
                    full_match_text = match.group(0).strip()
                    display_title = full_match_text if len(full_match_text) < 55 else f"{full_match_text[:52]}..."
                    self.log_print(f"   Logged Structural Section: '{display_title}' on Page {page_num + 1}")

            self.log_print(f"\n[Verification Engine Run Completed]")
            doc.save(final_pdf_path)
            doc.close()
            if os.path.exists(temp_pdf):
                os.remove(temp_pdf)
            self.log_print(f"\n[Finished] Output saved to:\n{final_pdf_path}")
            self.open_pdf_automatically(final_pdf_path)
            messagebox.showinfo("Success", "Analysis complete! Report compiled and exported.")

        except Exception as ex:
            self.log_print(f"\n❌ Execution Error encountered: {str(ex)}")
            messagebox.showerror("Execution Error", f"An error occurred: {str(ex)}")
            if "doc" in locals() and not doc.is_closed:
                doc.close()
            if os.path.exists(temp_pdf):
                os.remove(temp_pdf)
        finally:
            self.root.after(0, self.reset_ui_state)

    def reset_ui_state(self):
        self.run_btn.config(state="normal")
        self.browse_btn.config(state="normal")
        self.file_entry.config(state="normal")
        self.update_status("Ready", stop_progress=True)


if __name__ == "__main__":
    root = tk.Tk()
    app = DocVerifierApp(root)
    root.mainloop()