#!/usr/bin/env python3
"""
Etsy Packing Slip + Shipping Label Merger
-------------------------------------------
Merges Etsy packing slips and shipping labels into one PDF.

Layout modes:
  Half & Half  — label on top half, packing slip on bottom half,
                 both printed on ONE page per order. Cut in half to
                 get a ready-to-pack pair.
  Separate pages — original behavior: slip page then label page,
                   two pages per order.

Matching: each page is scanned for an Etsy order number and pages
sharing the same number are paired together. Works fully offline.
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
    from pypdf import PdfReader, PdfWriter, PageObject, Transformation
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

ORDER_NUMBER_PATTERNS = [
    re.compile(r'Order\s*#\s*[:\-]?\s*(\d{6,})', re.IGNORECASE),
    re.compile(r'Order\s*(?:No\.?|Number|ID)\s*[:#]?\s*(\d{6,})', re.IGNORECASE),
    re.compile(r'Order\s*[:#]\s*(\d{6,})', re.IGNORECASE),
    re.compile(r'Receipt\s*(?:ID|#)?\s*[:#]?\s*(\d{6,})', re.IGNORECASE),
    re.compile(r'Ref(?:erence)?\s*#?\s*[:\-]?\s*(\d{6,})', re.IGNORECASE),
]
FALLBACK_PATTERN = re.compile(r'\b(\d{9,10})\b')


def extract_order_number(text):
    if not text:
        return None
    for pattern in ORDER_NUMBER_PATTERNS:
        m = pattern.search(text)
        if m:
            return m.group(1)
    m = FALLBACK_PATTERN.search(text)
    return m.group(1) if m else None


def extract_pages_with_order_numbers(pdf_path, log):
    results = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            order_no = extract_order_number(text)
            results.append((i, order_no))
            log(f"    page {i + 1}: order number -> {order_no or 'NOT FOUND'}")
    return results


# ---------------------------------------------------------------------------
# Half-and-half layout
# ---------------------------------------------------------------------------

PAGE_W = 8.5 * 72   # 612 pts
PAGE_H = 11.0 * 72  # 792 pts
HALF_H = PAGE_H / 2  # 396 pts


def _fit_transform(src_page, dest_x, dest_y, dest_w, dest_h, padding=6):
    """Transformation that scales src_page to fit inside the dest box."""
    src_w = float(src_page.mediabox.width)
    src_h = float(src_page.mediabox.height)
    avail_w = dest_w - 2 * padding
    avail_h = dest_h - 2 * padding
    scale = min(avail_w / src_w, avail_h / src_h)
    x = dest_x + padding + (avail_w - src_w * scale) / 2
    y = dest_y + padding + (avail_h - src_h * scale) / 2
    return Transformation().scale(scale, scale).translate(x, y)


def make_combined_page(label_page, slip_page):
    """Letter-size page: label on top half, packing slip on bottom half."""
    new_page = PageObject.create_blank_page(width=PAGE_W, height=PAGE_H)
    new_page.merge_transformed_page(label_page, _fit_transform(label_page, 0, HALF_H, PAGE_W, HALF_H))
    new_page.merge_transformed_page(slip_page, _fit_transform(slip_page, 0, 0, PAGE_W, HALF_H))
    return new_page


# ---------------------------------------------------------------------------
# Merge logic
# ---------------------------------------------------------------------------

def merge_pdfs(slips_path, labels_path, output_path, layout, log):
    log("Reading packing slips PDF...")
    slip_pages = extract_pages_with_order_numbers(slips_path, log)

    log("")
    log("Reading shipping labels PDF...")
    label_pages = extract_pages_with_order_numbers(labels_path, log)

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
    log("Merging...")

    for slip_idx, order_no in slip_pages:
        s_page = slip_reader.pages[slip_idx]

        if order_no and label_map.get(order_no):
            matched_count += 1
            label_indices = label_map[order_no]

            if layout == "halfhalf":
                # First label + slip combined on one page
                l_page = label_reader.pages[label_indices[0]]
                used_label_indices.add(label_indices[0])
                writer.add_page(make_combined_page(l_page, s_page))
                log(f"    order {order_no}: label pg {label_indices[0]+1} + slip pg {slip_idx+1} -> combined")
                # Extra labels (multi-package) as separate pages
                for label_idx in label_indices[1:]:
                    writer.add_page(label_reader.pages[label_idx])
                    used_label_indices.add(label_idx)
                    log(f"    order {order_no}: extra label pg {label_idx+1} -> separate page")
            else:
                # Separate mode: slip then all labels
                writer.add_page(s_page)
                for label_idx in label_indices:
                    writer.add_page(label_reader.pages[label_idx])
                    used_label_indices.add(label_idx)
                log(f"    order {order_no}: slip pg {slip_idx+1} + label(s) {[i+1 for i in label_indices]}")
        else:
            writer.add_page(s_page)
            unmatched_slips.append(slip_idx)
            log(f"    slip pg {slip_idx+1} (order {order_no or 'unknown'}): NO MATCHING LABEL")

    # Leftover unmatched labels
    unmatched_labels = [idx for idx, _ in label_pages if idx not in used_label_indices]
    if unmatched_labels:
        log(f"\nAppending {len(unmatched_labels)} unmatched label(s) at end:")
        for idx in unmatched_labels:
            writer.add_page(label_reader.pages[idx])
            log(f"    label pg {idx+1}")

    with open(output_path, "wb") as f:
        writer.write(f)

    log(f"\nDone. {matched_count} of {len(slip_pages)} orders matched.")
    if unmatched_slips:
        log(f"WARNING: {len(unmatched_slips)} slip(s) had no matching label.")
    if unmatched_labels:
        log(f"WARNING: {len(unmatched_labels)} label(s) had no matching slip — appended at end.")

    return matched_count, len(unmatched_slips), len(unmatched_labels)


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class MergeApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Etsy Slip + Label Merger")
        self.geometry("720x620")
        self.minsize(640, 540)

        self.slips_path = tk.StringVar()
        self.labels_path = tk.StringVar()
        self.layout_var = tk.StringVar(value="halfhalf")

        pad = {"padx": 10, "pady": 5}

        tk.Label(self, text="Etsy Packing Slip + Shipping Label Merger",
                 font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=10, pady=(10, 2))

        # Layout choice
        lf = tk.LabelFrame(self, text="Layout", padx=8, pady=4)
        lf.pack(fill="x", padx=10, pady=(0, 6))
        tk.Radiobutton(lf, text="Half & Half — label on top, slip on bottom (one page per order)",
                       variable=self.layout_var, value="halfhalf",
                       font=("Segoe UI", 10, "bold")).pack(anchor="w")
        tk.Radiobutton(lf, text="Separate pages — slip page then label page",
                       variable=self.layout_var, value="separate",
                       font=("Segoe UI", 10)).pack(anchor="w")

        # File pickers
        row1 = tk.Frame(self)
        row1.pack(fill="x", **pad)
        tk.Label(row1, text="Packing Slips PDF:", width=20, anchor="w").pack(side="left")
        tk.Entry(row1, textvariable=self.slips_path).pack(side="left", fill="x", expand=True, padx=5)
        tk.Button(row1, text="Browse...", command=self.browse_slips).pack(side="left")

        row2 = tk.Frame(self)
        row2.pack(fill="x", **pad)
        tk.Label(row2, text="Shipping Labels PDF:", width=20, anchor="w").pack(side="left")
        tk.Entry(row2, textvariable=self.labels_path).pack(side="left", fill="x", expand=True, padx=5)
        tk.Button(row2, text="Browse...", command=self.browse_labels).pack(side="left")

        # Buttons
        row3 = tk.Frame(self)
        row3.pack(fill="x", **pad)
        self.merge_btn = tk.Button(row3, text="Merge", font=("Segoe UI", 11, "bold"),
                                    bg="#2e7d32", fg="white", command=self.run_merge)
        self.merge_btn.pack(side="left")
        self.open_btn = tk.Button(row3, text="Open Output Folder",
                                   command=self.open_output_folder, state="disabled")
        self.open_btn.pack(side="left", padx=8)

        # Log
        tk.Label(self, text="Log:").pack(anchor="w", padx=10)
        self.log_box = scrolledtext.ScrolledText(self, height=18, font=("Consolas", 9))
        self.log_box.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.output_dir = None

    def browse_slips(self):
        p = filedialog.askopenfilename(title="Select Packing Slips PDF",
                                        filetypes=[("PDF files", "*.pdf")])
        if p:
            self.slips_path.set(p)

    def browse_labels(self):
        p = filedialog.askopenfilename(title="Select Shipping Labels PDF",
                                        filetypes=[("PDF files", "*.pdf")])
        if p:
            self.labels_path.set(p)

    def log(self, msg):
        self.log_box.insert("end", msg + "\n")
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
        layout = self.layout_var.get()
        threading.Thread(target=self._worker, args=(slips, labels, layout), daemon=True).start()

    def _worker(self, slips, labels, layout):
        try:
            out_dir = os.path.dirname(slips)
            ts = datetime.now().strftime("%Y-%m-%d_%H%M%S")
            out_path = os.path.join(out_dir, f"Merged_{ts}.pdf")

            matched, us, ul = merge_pdfs(slips, labels, out_path, layout, self.log)
            self.output_dir = out_dir
            self.log(f"\nSaved: {out_path}")
            self.after(0, lambda: self.open_btn.config(state="normal"))

            if us or ul:
                self.after(0, lambda: messagebox.showwarning(
                    "Done (with warnings)",
                    f"{matched} order(s) matched.\n"
                    f"{us} slip(s) had no label.\n"
                    f"{ul} label(s) had no slip.\n\nCheck log before printing."
                ))
            else:
                self.after(0, lambda: messagebox.showinfo("Done",
                    f"All {matched} orders merged successfully."))
        except Exception:
            self.log("\nERROR:\n" + traceback.format_exc())
            self.after(0, lambda: messagebox.showerror("Error", "See log for details."))
        finally:
            self.after(0, lambda: self.merge_btn.config(state="normal", text="Merge"))

    def open_output_folder(self):
        if self.output_dir and os.path.isdir(self.output_dir):
            try:
                os.startfile(self.output_dir)
            except AttributeError:
                messagebox.showinfo("Output folder", self.output_dir)


if __name__ == "__main__":
    app = MergeApp()
    app.mainloop()
