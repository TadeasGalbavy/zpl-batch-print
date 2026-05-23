import os
import re
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

# ---- Voliteľné: win32print pre Windows/USB tlač ----
WIN32_OK = True
try:
    import win32print
except Exception:
    WIN32_OK = False

APP_TITLE = "Tlač štítkov"
# názov súboru
DEFAULT_SHEET = "TLAČ"
# cesta k súborom
FIXED_PRN_DIR = r"G:\.shortcut-targets-by-id\1cw7UvlbZhSWGjBJus8DT_b2XDJAhFeUJ\TLAC_STITKOV\~PRN"

PQ_REGEX = re.compile(r"\^PQ[^\^]*")  # nájde existujúci ^PQ

def set_quantity(zpl_text: str, qty: int) -> str:
    """Vloží/aktualizuje ^PQ príkaz. Ak neexistuje, vloží pred ^XZ."""
    new_pq = f"^PQ{qty},0,1,N"  # qty, pause=0, rep=1, no reprint
    if PQ_REGEX.search(zpl_text):
        return PQ_REGEX.sub(new_pq, zpl_text)
    if "^XZ" in zpl_text:
        return zpl_text.replace("^XZ", f"{new_pq}^XZ")
    # fallback: pridaj na koniec, ak by náhodou ^XZ chýbal
    return zpl_text + new_pq

def load_prn(path: str) -> str:
    with open(path, "r", encoding="latin-1") as f:
        return f.read()

def send_zpl_windows(zpl_bytes: bytes, printer_name: str):
    if not WIN32_OK:
        raise RuntimeError("win32print nie je dostupný. Nainštaluj balík 'pywin32'.")
    hPrinter = win32print.OpenPrinter(printer_name)
    try:
        job = win32print.StartDocPrinter(hPrinter, 1, ("Auto ZPL print", None, "RAW"))
        win32print.StartPagePrinter(hPrinter)
        win32print.WritePrinter(hPrinter, zpl_bytes)
        win32print.EndPagePrinter(hPrinter)
        win32print.EndDocPrinter(hPrinter)
    finally:
        win32print.ClosePrinter(hPrinter)

def read_joblist_from_excel(xlsx_path: str):
    df = pd.read_excel(xlsx_path)

    # odstráň prázdne riadky
    df = df.dropna(how="all")

    # kontrola minimálne 2 stĺpcov
    if df.shape[1] < 2:
        raise ValueError("Excel musí obsahovať aspoň 2 stĺpce (súbor, množstvo).")

    rows = []
    for _, r in df.iterrows():
        file_name = str(r.iloc[0]).strip()   # 1. stĺpec
        try:
            qty = int(r.iloc[1])            # 2. stĺpec
        except Exception:
            raise ValueError(f"Neplatná hodnota množstva pri súbore: {file_name}")

        rows.append({
            "sablona": file_name,
            "mnozstvo": qty
        })

    return rows

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("780x540")
        self.minsize(760, 520)

        # Vars
        self.excel_path = tk.StringVar()
        self.sheet_name = tk.StringVar(value=DEFAULT_SHEET)
        self.prn_dir = tk.StringVar()
        self.excel_path = tk.StringVar(value="")
        self.excel_name = tk.StringVar(value="Vybraný súbor: (žiadny)")
        self.prn_dir.set(FIXED_PRN_DIR)
        self.printer_name = tk.StringVar()
        self.simulate = tk.BooleanVar(value=False)

        # UI
        self._build_ui()

        # Načítaj zoznam tlačiarní (ak vieme)
        if WIN32_OK:
            self._refresh_printers()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        # frm_top = ttk.LabelFrame(self, text="Vstupy")
        # frm_top.pack(fill="x", **pad)

        # Excel
        # row1 = ttk.Frame(frm_top)
        # row1.pack(fill="x", padx=10, pady=(10, 4))
        # ttk.Label(row1, text="Excel súbor:").pack(side="left")
        # ttk.Entry(row1, textvariable=self.excel_path, width=60).pack(side="left", padx=8)
        # ttk.Button(row1, text="Vybrať…", command=self._pick_excel).pack(side="left")
        # ttk.Label(row1, text="List:").pack(side="left", padx=(16, 4))
        # ttk.Entry(row1, textvariable=self.sheet_name, width=16).pack(side="left")

        # PRN dir
        # row2 = ttk.Frame(frm_top)
        # row2.pack(fill="x", padx=10, pady=4)
        # ttk.Label(row2, text="Zložka s .prn:").pack(side="left")
        # ttk.Entry(row2, textvariable=self.prn_dir, width=60).pack(side="left", padx=8)
        # ttk.Button(row2, text="Vybrať…", command=self._pick_prn_dir).pack(side="left")

        # Mode
        frm_mode = ttk.LabelFrame(self, text="Tlačiareň (Windows)")
        frm_mode.pack(fill="x", **pad)

        col1 = ttk.Frame(frm_mode)
        col1.pack(fill="x", padx=10, pady=8)

        ttk.Label(col1, text="Vyber tlačiareň:").grid(row=0, column=0, sticky="w")
        self.cmb_printers = ttk.Combobox(col1, textvariable=self.printer_name, width=50, state="readonly")
        self.cmb_printers.grid(row=0, column=1, sticky="w", padx=(10, 0))
        ttk.Button(col1, text="Obnoviť zoznam", command=self._refresh_printers).grid(row=0, column=2, padx=12)
        ttk.Button(col1, text="Vybrať súbor", command=self._pick_excel).grid(row=0, column=3, padx=12)
        self.lbl_excel = ttk.Label(col1, textvariable=self.excel_name, foreground="red")
        self.lbl_excel.grid(row=1, column=1, columnspan=3, sticky="w", padx=(10, 0), pady=(8, 0))

        # Options
        frm_opts = ttk.Frame(self)
        frm_opts.pack(fill="x", padx=10, pady=4)
        ttk.Checkbutton(frm_opts, text="Simulácia (neposielať do tlačiarne)", variable=self.simulate).pack(side="left")

        # Buttons
        frm_btns = ttk.Frame(self)
        frm_btns.pack(fill="x", padx=10, pady=6)
        ttk.Button(frm_btns, text="Spustiť tlač", command=self._start_print).pack(side="left")
        ttk.Button(frm_btns, text="Koniec", command=self.destroy).pack(side="right")

        # Log
        frm_log = ttk.LabelFrame(self, text="Log")
        frm_log.pack(fill="both", expand=True, padx=10, pady=(6, 10))
        self.txt_log = tk.Text(frm_log, height=12, wrap="word")
        self.txt_log.pack(fill="both", expand=True, padx=8, pady=8)

    def _pick_excel(self):
        path = filedialog.askopenfilename(
            title="Vyber Excel",
            filetypes=[("Excel súbor", "*.xlsx *.xls"), ("Všetky súbory", "*.*")]
        )
        if path:
            self.excel_path.set(path)  # full path pre logiku
            self.excel_name.set(f"Vybraný súbor: {os.path.basename(path)}")  # len názov pre UI
            self.lbl_excel.config(foreground="green")
            self._log(f"Vybraný Excel: {path}")

    def _pick_prn_dir(self):
        path = filedialog.askdirectory(title="Vyber zložku s .prn")
        if path:
            self.prn_dir.set(path)

    def _log(self, msg: str):
        self.txt_log.insert("end", msg + "\n")
        self.txt_log.see("end")
        self.update_idletasks()

    def _refresh_printers(self):
        if not WIN32_OK:
            self.cmb_printers["values"] = []
            self.printer_name.set("")
            self._log("Upozornenie: 'pywin32' nie je nainštalované – režim Windows/USB nebude fungovať.")
            return
        try:
            flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            printers = win32print.EnumPrinters(flags)
            names = [p[2] for p in printers if "zdesigner" in p[2].lower()]  # p[2] = názov
            self.cmb_printers["values"] = names
            # predvyber predvolenú
            try:
                default_name = win32print.GetDefaultPrinter()
                if default_name in names:
                    self.printer_name.set(default_name)
                elif names:
                    self.printer_name.set(names[0])
            except Exception:
                if names:
                    self.printer_name.set(names[0])
            self._log(f"Nájdené tlačiarne: {', '.join(names) if names else '(žiadne)'}")
        except Exception as e:
            self._log(f"CHYBA pri načítaní tlačiarní: {e}")

    def _start_print(self):
        # spusti na pozadí, nech GUI nezamrzne
        threading.Thread(target=self._do_print, daemon=True).start()

    def _do_print(self):
        try:
            xlsx = self.excel_path.get().strip()
            prn_dir = FIXED_PRN_DIR

            if not xlsx:
                messagebox.showerror(APP_TITLE, "Prosím vyber Excel súbor.")
                return

            if not os.path.isfile(xlsx):
                messagebox.showerror(APP_TITLE, "Vybraný Excel súbor neexistuje.")
                return

            self._log(f"Načítavam Excel: {xlsx}.")
            jobs = read_joblist_from_excel(xlsx)

            self._log(f"Počet riadkov na spracovanie: {len(jobs)}")

            total_labels = sum(row["mnozstvo"] for row in jobs)
            self._log(f"Počet štítkov na tlač: {total_labels}")

            simulate = self.simulate.get()

            if not WIN32_OK:
                raise RuntimeError("Režim Windows/USB vyžaduje balík 'pywin32' (pip install pywin32).")

            printer = self.printer_name.get().strip()
            if not printer:
                raise ValueError("Vyber Windows tlačiareň.")

            self._log(f"Režim: WINDOWS/USB – {printer}")

            processed = 0
            skipped = 0
            for row in jobs:
                file_name = row["sablona"]
                qty = row["mnozstvo"]
                prn_path = os.path.join(prn_dir, file_name)
                if not os.path.isfile(prn_path):
                    self._log(f"SKIP: Nenájdené: {prn_path}")
                    skipped += 1
                    continue
                try:
                    zpl = load_prn(prn_path)
                    zpl = set_quantity(zpl, qty)
                    data = zpl.encode("latin-1", errors="ignore")

                    if simulate:
                        self._log(f"[SIMULÁCIA] {file_name} × {qty}")
                    else:
                        send_zpl_windows(data, printer)
                        self._log(f"OK: {file_name} × {qty}")
                    processed += 1
                except Exception as e:
                    self._log(f"CHYBA pri {file_name}: {e}")
                    self._log(traceback.format_exc())

            self._log("—" * 60)
            self._log(f"Hotovo. Spracované: {processed}, preskočené: {skipped}, riadkov: {len(jobs)}, štítkov: {total_labels}")
            messagebox.showinfo(APP_TITLE, "Tlač dokončená – pozri log.")
        except Exception as e:
            self._log(f"FATAL: {e}")
            self._log(traceback.format_exc())
            messagebox.showerror(APP_TITLE, f"Chyba: {e}")

if __name__ == "__main__":
    App().mainloop()