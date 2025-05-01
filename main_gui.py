"""
GA Broker • Excel Splitter   (CustomTkinter GUI, no progress bar)
-----------------------------------------------------------------
• Lets user pick source Excel and rows-per-file
• Calls split_excel_core.save_chunks(...) in a background thread
• Saves output to /splitted_excels
"""

import os, sys, threading, customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image
import split_excel_core as backend

APP_DIR = os.path.dirname(sys.executable if getattr(sys, "frozen", False) else __file__)
BANNER  = os.path.join(APP_DIR, "company_banner.png")
ICON    = os.path.join(APP_DIR, "company_logo.ico")
OUT_DIR = os.path.join(APP_DIR, "splitted_excels")

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class SplitterUI(ctk.CTk):
    def __init__(self):
        super().__init__()

        # ── window / icon ──
        self.title("GA Broker • Excel Splitter")
        if os.path.exists(ICON):
            try:
                self.iconbitmap(ICON)
            except Exception:
                pass

        # ── banner ──
        banner_img = Image.open(BANNER)
        bw, bh = banner_img.size
        ctk.CTkLabel(
            self,
            image=ctk.CTkImage(light_image=banner_img, dark_image=banner_img, size=(bw, bh)),
            text="",
        ).pack(pady=(10, 20))

        self.geometry(f"{bw + 40}x300")
        self.resizable(False, False)

        # ── file picker row ──
        row = ctk.CTkFrame(self, fg_color="transparent")
        row.pack(pady=5, fill="x", padx=20)

        ctk.CTkLabel(row, text="Source Excel:").pack(side="left", padx=(0, 5))
        self.path_var = ctk.StringVar()
        ctk.CTkEntry(row, width=bw - 180, textvariable=self.path_var).pack(side="left", padx=5)
        ctk.CTkButton(row, text="Browse…", command=self.browse).pack(side="left")

        # ── rows-per-file row ──
        row2 = ctk.CTkFrame(self, fg_color="transparent")
        row2.pack(pady=2, fill="x", padx=20)
        ctk.CTkLabel(row2, text="Rows per file:").pack(side="left", padx=(0, 5))
        self.rows_var = ctk.StringVar(value="998")
        ctk.CTkEntry(row2, width=120, textvariable=self.rows_var).pack(side="left")

        # ── buttons ──
        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.pack(pady=15)
        self.run_btn = ctk.CTkButton(btns, text="Run", width=120, command=self.run_clicked)
        self.run_btn.pack(side="left", padx=10)
        ctk.CTkButton(btns, text="Exit", width=120, command=self.destroy).pack(side="left", padx=10)

    # ── UI helpers ──
    def browse(self):
        file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file:
            self.path_var.set(file)

    def run_clicked(self):
        src = self.path_var.get()
        if not src:
            messagebox.showerror("No file selected", "Please pick an Excel file.")
            return

        try:
            rows = int(self.rows_var.get())
            if rows <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Invalid number", "Rows per file must be a positive integer.")
            return

        self.run_btn.configure(state="disabled")
        threading.Thread(target=self._worker, args=(src, rows), daemon=True).start()

    def _worker(self, src_path, rows_per_file):
        try:
            mawb = backend.get_mawb(src_path)
            df = backend.build_dataframe(src_path)
            parts = backend.save_chunks(df, OUT_DIR, mawb, rows_per_file)
            self.after(0, lambda p=parts: self._done(p))
        except Exception as exc:
            self.after(0, lambda err=exc: self._error(err))

    def _done(self, parts):
        self.run_btn.configure(state="normal")
        messagebox.showinfo("Done", f"Finished – {parts} file(s) saved to:\n{OUT_DIR}")
        try:
            os.startfile(OUT_DIR)
        except Exception:
            pass

    def _error(self, err):
        self.run_btn.configure(state="normal")
        messagebox.showerror("Error", str(err))


if __name__ == "__main__":
    os.makedirs(OUT_DIR, exist_ok=True)
    SplitterUI().mainloop()
