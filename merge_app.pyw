#!/usr/bin/env python3
"""
Etsy Packing Slip + Shipping Label Merger
-------------------------------------------
Offline desktop tool that merges an Etsy "Packing Slips" PDF and an Etsy
"Shipping Labels" PDF into one PDF, with each packing slip immediately
followed by its matching shipping label(s) -- so when you bulk-print the
merged file once, you get slip+label pairs in order, ready to pack.

How matching works:
  Each page of both PDFs is scanned for an Etsy order number (a long
  numeric ID, usually shown as "Order #1234567890" or similar). Slip pages
  and label pages that share the same order number are paired together.

No internet connection is required to run this after the one-time setup
described in README.txt.
"""

import os
import re
import sys
import threading
import traceback
from collections import defaultdict
from datetime import datetime

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

try:
    import pdfplumber
    from pypdf import PdfReader, PdfWriter
except ImportError:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        "Missing dependencies",
        "Required packages are not installed.\n\n"
        "Please run install.bat first (see README.txt), then re-open this app.",
    )
    sys.exit(1)


# ---------------------------------------------------------------------------
# Order-number extraction
# ---------------------------------------------------------------------------

# Tried in order, most specific first. All require a run of 6+ digits so we
# don't accidentally grab a zip code or a short quantity number.
ORDER_NUMBER_PATTERNS = [
    re.compile(r'Order\s*#\s*[:\-]?\s*(\d{6,})', re.IGNORECASE),
    re.compile(r'Order\s*(?:No\.?|Number|ID)\s*[:#]?\s*(\d{6,})', re.IGNORECASE),
    re.compile(r'Order\s*[:#]\s*(\d{6,})', re.IGNORECASE),
    re.compile(r'Receipt\s*(?:ID|#)?\s*[:#]?\s*(\d{6,})', re.IGNORECASE),
    re.compile(r'Ref(?:erence)?\s*#?\s*[:\-]?\s*(\d{6,})', re.IGNORECASE),
]

# Pure fallback: a lone 9-10 digit number with nothing else on context.
FALLBACK_PATTERN = re.compile(r'\b(\d{9,10})\b')


def extract_order_number(text):
    """Return the first order number found in a page's text, or None."""
    if not text:
        return None
    for pattern in ORDER_NUMBER_PATTERNS:
        m = pattern.search(text)
        if m:
            return m.group(1)
    # Fallback: grab the first 9-10 digit standalone number (Etsy order IDs
    # are typically this length). Less reliable, used only if nothing else hit.
    m = FALLBACK_PATTERN.search(text)
    if m:
        return m.group(1)
    return None


def extract_pages_with_order_numbers(pdf_path, log):
    """Returns list of (page_index, order_number_or_None) for every page."""
    results = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            order_no = extract_order_number(text)
            results.append((i, order_no))
            label = order_no if order_no else "NOT FOUND"
            log(f"    page {i + 1}: order number -> {label}")
    return results


# ---------------------------------------------------------------------------
# Merge logic
# ---------------------------------------------------------------------------

def merge_pdfs(slips_path, labels_path, output_path, log):
    log("Reading packing slips PDF...")
    slip_pages = extract_pages_with_order_numbers(slips_path, log)

    log("")
    log("Reading shipping labels PDF...")
    label_pages = extract_pages_with_order_numbers(labels_path, log)

    # Map order number -> list of label page indices (an order can have
    # more than one label/package).
    label_map = defaultdict(list)
    for idx, order_no in label_pages:
        if order_no:
            label_map[order_no].append(idx)

    slip_reader = PdfReader(slips_path)
    label_reader = PdfReader(labels_path)
    writer = PdfWriter()

    used_label_indices = set()
    unmatched_slips = []
    matched_count = 0

    log("")
    log("Matching slips to labels...")
    for slip_idx, order_no in slip_pages:
        writer.add_page(slip_reader.pages[slip_idx])
        if order_no and label_map.get(order_no):
            matched_count += 1
            for label_idx in label_map[order_no]:
                writer.add_page(label_reader.pages[label_idx])
                used_label_indices.add(label_idx)
            log(f"    slip page {slip_idx + 1} (order {order_no}) "
                f"<-> label page(s) {[i + 1 for i in label_map[order_no]]}")
        else:
            unmatched_slips.append(slip_idx)
            log(f"    slip page {slip_idx + 1} (order {order_no or 'unknown'}): "
                f"NO MATCHING LABEL FOUND")

    # Any label pages that were never used get appended at the end so
    # nothing is silently dropped.
    unmatched_labels = [idx for idx, _ in label_pages if idx not in used_label_indices]
    if unmatched_labels:
        log("")
        log(f"Appending {len(unmatched_labels)} unmatched label page(s) at the end:")
        for idx in unmatched_labels:
            writer.add_page(label_reader.pages[idx])
            log(f"    label page {idx + 1} appended (no matching slip found)")

    with open(output_path, "wb") as f:
        writer.write(f)

    log("")
    log(f"Done. {matched_count} of {len(slip_pages)} slips matched with a label.")
    if unmatched_slips:
        log(f"WARNING: {len(unmatched_slips)} slip page(s) had no matching label "
            f"-- check the log above for page numbers, and verify manually before printing.")
    if unmatched_labels:
        log(f"WARNING: {len(unmatched_labels)} label page(s) had no matching slip "
            f"-- they were appended at the end of the merged file.")

    return matched_count, len(unmatched_slips), len(unmatched_labels)


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class MergeApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Etsy Packing Slip + Shipping Label Merger")
        self.geometry("720x560")
        self.minsize(640, 480)

        self.slips_path = tk.StringVar()
        self.labels_path = tk.StringVar()

        pad = {"padx": 10, "pady": 6}

        tk.Label(self, text="Etsy Packing Slip + Shipping Label Merger",
                 font=("Segoe UI", 14, "bold")).pack(anchor="w", **pad)
        tk.Label(self, text="Merges your packing slips PDF and shipping labels PDF, "
                             "pairing each slip with its matching label by order number.",
                 wraplength=680, justify="left", fg="#444").pack(anchor="w", padx=10)

        # Packing slips row
        row1 = tk.Frame(self)
        row1.pack(fill="x", **pad)
        tk.Label(row1, text="Packing Slips PDF:", width=20, anchor="w").pack(side="left")
        tk.Entry(row1, textvariable=self.slips_path).pack(side="left", fill="x", expand=True, padx=5)
        tk.Button(row1, text="Browse...", command=self.browse_slips).pack(side="left")

        # Shipping labels row
        row2 = tk.Frame(self)
        row2.pack(fill="x", **pad)
        tk.Label(row2, text="Shipping Labels PDF:", width=20, anchor="w").pack(side="left")
        tk.Entry(row2, textvariable=self.labels_path).pack(side="left", fill="x", expand=True, padx=5)
        tk.Button(row2, text="Browse...", command=self.browse_labels).pack(side="left")

        # Merge button
        row3 = tk.Frame(self)
        row3.pack(fill="x", **pad)
        self.merge_btn = tk.Button(row3, text="Merge", font=("Segoe UI", 11, "bold"),
                                    bg="#2e7d32", fg="white", command=self.run_merge)
        self.merge_btn.pack(side="left")
        self.open_folder_btn = tk.Button(row3, text="Open Output Folder",
                                          command=self.open_output_folder, state="disabled")
        self.open_folder_btn.pack(side="left", padx=8)

        # Log box
        tk.Label(self, text="Log:").pack(anchor="w", padx=10)
        self.log_box = scrolledtext.ScrolledText(self, height=20, font=("Consolas", 9))
        self.log_box.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.output_dir = None

    def browse_slips(self):
        path = filedialog.askopenfilename(title="Select Packing Slips PDF",
                                           filetypes=[("PDF files", "*.pdf")])
        if path:
            self.slips_path.set(path)

    def browse_labels(self):
        path = filedialog.askopenfilename(title="Select Shipping Labels PDF",
                                           filetypes=[("PDF files", "*.pdf")])
        if path:
            self.labels_path.set(path)

    def log(self, message):
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")
        self.update_idletasks()

    def run_merge(self):
        slips = self.slips_path.get().strip()
        labels = self.labels_path.get().strip()

        if not slips or not os.path.isfile(slips):
            messagebox.showerror("Missing file", "Please select a valid Packing Slips PDF.")
            return
        if not labels or not os.path.isfile(labels):
            messagebox.showerror("Missing file", "Please select a valid Shipping Labels PDF.")
            return

        self.log_box.delete("1.0", "end")
        self.merge_btn.config(state="disabled", text="Merging...")

        thread = threading.Thread(target=self._merge_worker, args=(slips, labels), daemon=True)
        thread.start()

    def _merge_worker(self, slips, labels):
        try:
            out_dir = os.path.dirname(slips)
            timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
            output_path = os.path.join(out_dir, f"Merged_Slips_Labels_{timestamp}.pdf")

            matched, unmatched_slips, unmatched_labels = merge_pdfs(
                slips, labels, output_path, self.log
            )

            self.output_dir = out_dir
            self.log("")
            self.log(f"Saved merged PDF to:\n{output_path}")

            self.after(0, lambda: self.open_folder_btn.config(state="normal"))

            if unmatched_slips or unmatched_labels:
                self.after(0, lambda: messagebox.showwarning(
                    "Merge complete (with warnings)",
                    f"Merged PDF saved.\n\n"
                    f"{matched} slip(s) matched.\n"
                    f"{unmatched_slips} slip(s) had no matching label.\n"
                    f"{unmatched_labels} label(s) had no matching slip.\n\n"
                    f"Check the log and verify the output before printing."
                ))
            else:
                self.after(0, lambda: messagebox.showinfo(
                    "Merge complete", f"All {matched} slips matched successfully."
                ))
        except Exception:
            err = traceback.format_exc()
            self.log("")
            self.log("ERROR:\n" + err)
            self.after(0, lambda: messagebox.showerror(
                "Error", "Something went wrong during merge. See the log for details."
            ))
        finally:
            self.after(0, lambda: self.merge_btn.config(state="normal", text="Merge"))

    def open_output_folder(self):
        if self.output_dir and os.path.isdir(self.output_dir):
            try:
                os.startfile(self.output_dir)  # Windows
            except AttributeError:
                messagebox.showinfo("Output folder", self.output_dir)


if __name__ == "__main__":
    app = MergeApp()
    app.mainloop()
