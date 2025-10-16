#!/usr/bin/env python3
"""
Excel Compare GUI
- Select two Excel files
- Enter ranges (e.g. A1:A100)
- Choose target columns for matched pairs (e.g. C and D)
- Compare, color green=match, red=different, and write matches side-by-side
- Saves: <file1>_colored.xlsx, <file2>_colored.xlsx, result_matches.xlsx
Requires: openpyxl
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
import os
import re

GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

def parse_range(range_str):
    """
    Accepts Excel-like ranges like 'A1:A100' or 'B2:B10' or 'A:A' or 'C1:C'.
    Returns (sheet_range_col, start_row, end_row) where end_row may be None to mean until last used row.
    """
    if not range_str:
        return None
    rng = range_str.replace(" ", "").upper()
    # Simple pattern: COLROW:COLROW or COL:COL
    m = re.match(r'^([A-Z]+)(\d*):([A-Z]+)(\d*)$', rng)
    if m:
        col1, r1, col2, r2 = m.groups()
        start = int(r1) if r1 else 1
        end = int(r2) if r2 else None
        return (col1, start, end)
    m2 = re.match(r'^([A-Z]+):([A-Z]+)$', rng)
    if m2:
        col1, col2 = m2.groups()
        return (col1, 1, None)
    raise ValueError("Aralık biçimi geçersiz. Örnek: A1:A100 veya B2:B200 veya A:A")

def get_column_index(col_letters):
    idx = 0
    for ch in col_letters:
        idx = idx*26 + (ord(ch)-64)
    return idx

def read_values_from_workbook(path, range_str):
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    col, start, end = parse_range(range_str)
    col_idx = get_column_index(col)
    max_row = ws.max_row
    if end is None:
        end = max_row
    values = []
    coords = []
    for r in range(start, end+1):
        cell = ws.cell(row=r, column=col_idx)
        values.append(cell.value if cell.value is not None else "")
        coords.append((r, col_idx))
    return values, coords, wb, ws

def color_and_save(wb, ws, is_match_map, out_path):
    """
    is_match_map: dict with keys (row, col) -> True/False
    Colors cells in ws according to map and saves workbook to out_path.
    """
    for (r, c), is_match in is_match_map.items():
        cell = ws.cell(row=r, column=c)
        try:
            cell.fill = GREEN_FILL if is_match else RED_FILL
        except Exception:
            # ignore cells that can't be filled for any reason
            pass
    wb.save(out_path)

def compare_and_write(file1, range1, file2, range2, target_col1, target_col2, case_sensitive=False):
    vals1, coords1, wb1, ws1 = read_values_from_workbook(file1, range1)
    vals2, coords2, wb2, ws2 = read_values_from_workbook(file2, range2)
    # prepare lookup for vals2
    lookup2 = {}
    for v in vals2:
        key = v if case_sensitive else (str(v).lower() if v is not None else "")
        lookup2[key] = lookup2.get(key, 0) + 1

    is_match_map1 = {}
    is_match_map2 = {}
    matches = []

    for i, v in enumerate(vals1):
        key = v if case_sensitive else (str(v).lower() if v is not None else "")
        match = False
        if key in lookup2 and lookup2[key] > 0 and key != "":
            match = True
            # reduce count to avoid duplicate multiple matches mapping
            lookup2[key] -= 1
        is_match_map1[coords1[i]] = match
        if match:
            matches.append(v)

    # For second file, determine matches by cross-checking against original vals1 list
    lookup1 = {}
    for v in vals1:
        key = v if case_sensitive else (str(v).lower() if v is not None else "")
        lookup1[key] = lookup1.get(key, 0) + 1

    for i, v in enumerate(vals2):
        key = v if case_sensitive else (str(v).lower() if v is not None else "")
        match = False
        if key in lookup1 and lookup1[key] > 0 and key != "":
            match = True
            lookup1[key] -= 1
        is_match_map2[coords2[i]] = match

    # Color and save copies
    base1 = os.path.splitext(os.path.basename(file1))[0]
    base2 = os.path.splitext(os.path.basename(file2))[0]
    out1 = os.path.join(os.getcwd(), f"{base1}_colored.xlsx")
    out2 = os.path.join(os.getcwd(), f"{base2}_colored.xlsx")
    color_and_save(wb1, ws1, is_match_map1, out1)
    color_and_save(wb2, ws2, is_match_map2, out2)

    # Write matches into result workbook side-by-side
    result_wb = Workbook()
    result_ws = result_wb.active
    result_ws.title = "Matches"
    # Headers
    result_ws.cell(row=1, column=get_column_index(target_col1)).value = f"From {base1}"
    result_ws.cell(row=1, column=get_column_index(target_col2)).value = f"From {base2}"
    row_counter = 2
    # We will iterate through original lists and record where both marked as match.
    # For simplicity, take matches from vals1 in order and fill both cols with same value.
    for v in matches:
        result_ws.cell(row=row_counter, column=get_column_index(target_col1)).value = v
        result_ws.cell(row=row_counter, column=get_column_index(target_col2)).value = v
        row_counter += 1

    result_path = os.path.join(os.getcwd(), "result_matches.xlsx")
    result_wb.save(result_path)
    return out1, out2, result_path

# ---- GUI ----
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel Compare — Renkli Arayüz")
        self.geometry("700x420")
        self.resizable(False, False)
        self.configure(bg="#f4f7fb")

        style = ttk.Style(self)
        style.theme_use('clam')

        # Big frame
        frm = ttk.Frame(self, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        lbl_title = tk.Label(frm, text="Excel Karşılaştırma Aracı", font=("Segoe UI", 16, "bold"), bg="#f4f7fb")
        lbl_title.grid(row=0, column=0, columnspan=3, pady=(0,10), sticky="w")

        # File selectors
        self.file1_var = tk.StringVar()
        self.file2_var = tk.StringVar()
        ttk.Button(frm, text="📁 Dosya 1 Seç", command=self.pick_file1).grid(row=1, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.file1_var, width=60).grid(row=1, column=1, columnspan=2, padx=6, sticky="w")

        ttk.Button(frm, text="📁 Dosya 2 Seç", command=self.pick_file2).grid(row=2, column=0, sticky="w", pady=6)
        ttk.Entry(frm, textvariable=self.file2_var, width=60).grid(row=2, column=1, columnspan=2, padx=6, sticky="w")

        ttk.Separator(frm, orient=tk.HORIZONTAL).grid(row=3, column=0, columnspan=3, sticky="ew", pady=10)

        # Range inputs
        self.range1_var = tk.StringVar(value="A1:A100")
        self.range2_var = tk.StringVar(value="A1:A200")
        ttk.Label(frm, text="1. Dosya Aralığı (örn: A1:A100):").grid(row=4, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.range1_var, width=20).grid(row=4, column=1, sticky="w")
        ttk.Label(frm, text="2. Dosya Aralığı (örn: A1:A200):").grid(row=5, column=0, sticky="w", pady=6)
        ttk.Entry(frm, textvariable=self.range2_var, width=20).grid(row=5, column=1, sticky="w", pady=6)

        # Target columns
        self.target_c_var = tk.StringVar(value="C")
        self.target_d_var = tk.StringVar(value="D")
        ttk.Label(frm, text="Aynıların yazılacağı sütun 1 (örn: C):").grid(row=6, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.target_c_var, width=6).grid(row=6, column=1, sticky="w")
        ttk.Label(frm, text="Aynıların yazılacağı sütun 2 (örn: D):").grid(row=7, column=0, sticky="w", pady=6)
        ttk.Entry(frm, textvariable=self.target_d_var, width=6).grid(row=7, column=1, sticky="w", pady=6)

        # Case sensitive checkbox
        self.case_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm, text="Büyük/Küçük harf duyarlılığı (case-sensitive)", variable=self.case_var).grid(row=8, column=0, columnspan=3, sticky="w", pady=8)

        # Compare button
        self.compare_btn = ttk.Button(frm, text="Karşılaştır ve Kaydet", command=self.run_compare)
        self.compare_btn.grid(row=9, column=0, columnspan=3, pady=14)

        # Status
        self.status_var = tk.StringVar(value="Hazır")
        ttk.Label(frm, textvariable=self.status_var).grid(row=10, column=0, columnspan=3, sticky="w")

    def pick_file1(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx;*.xlsm;*.xls")])
        if p:
            self.file1_var.set(p)

    def pick_file2(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx;*.xlsm;*.xls")])
        if p:
            self.file2_var.set(p)

    def run_compare(self):
        f1 = self.file1_var.get().strip()
        f2 = self.file2_var.get().strip()
        r1 = self.range1_var.get().strip()
        r2 = self.range2_var.get().strip()
        tc = self.target_c_var.get().strip().upper()
        td = self.target_d_var.get().strip().upper()
        cs = self.case_var.get()
        if not all([f1, f2, r1, r2, tc, td]):
            messagebox.showwarning("Eksik bilgi", "Lütfen tüm alanları doldurun.")
            return
        try:
            self.status_var.set("Karşılaştırılıyor...")
            out1, out2, res = compare_and_write(f1, r1, f2, r2, tc, td, case_sensitive=cs)
            self.status_var.set("Tamamlandı. Dosyalar kaydedildi.")
            messagebox.showinfo("Bitti", f"İşlem tamamlandı.\n\nKaydedilen dosyalar:\n{out1}\n{out2}\n{res}")
        except Exception as e:
            messagebox.showerror("Hata", f"Karşılaştırma sırasında hata oluştu:\n{e}")
            self.status_var.set("Hata oluştu.")

if __name__ == "__main__":
    App().mainloop()
