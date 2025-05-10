# main_ui.py
import os, sys, threading
from pathlib import Path
import customtkinter as ctk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
from PIL import Image


import mini_updater as updater
import excel_splitter as splitter
from reject_code_sorter import RejectCodeSorterTab
from pga_reference import PGAReferenceTab




# ─────────────────────────── paths ────────────────────────────
# When frozen, sys.executable = …\<ver>\_internal\GA Broker Helper.exe
APP_DIR  = Path(sys.executable if getattr(sys, "frozen", False)
                else __file__).resolve().parent      # → …\_internal
LOGO_DIR = APP_DIR / "Resources" / "Logo"
OUT_DIR  = APP_DIR / "splitted_excels"

BANNER = LOGO_DIR / "company_banner.png"
ICON   = LOGO_DIR / "company_logo.ico"


# ────────────────────── CTk global style ──────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


# ────────────────────────── tabs ──────────────────────────────
class ExcelSplitterTab:
    def __init__(self, parent):
        self.file_path = ""

        ctk.CTkLabel(
            parent, text="Drag & Drop Excel File Here or Use Browse",
            font=("Arial", 14), anchor="w", wraplength=600
        ).pack(padx=20, pady=(20, 10))

        # --- drop zone ---
        self.drop_target = ctk.CTkFrame(parent, height=60, width=400,
                                        fg_color="#808080", corner_radius=10)
        self.drop_target.pack(pady=5)
        self.drop_target.pack_propagate(False)

        self.drop_info = ctk.CTkLabel(self.drop_target, text="No file selected",
                                      font=("Arial", 12), text_color="#000000")
        self.drop_info.pack(expand=True)

        self.drop_target.drop_target_register(DND_FILES)
        self.drop_target.dnd_bind("<<Drop>>", self.on_drop)

        ctk.CTkButton(parent, text="Browse File",
                      command=self.browse).pack(pady=5)

        # --- rows-per-file ---
        row_frame = ctk.CTkFrame(parent, fg_color="transparent")
        row_frame.pack(pady=4)
        ctk.CTkLabel(row_frame, text="Rows per file:").pack(side="left", padx=(0, 5))
        self.rows_var = ctk.StringVar(value="499")
        ctk.CTkEntry(row_frame, width=80,
                     textvariable=self.rows_var).pack(side="left")

        # --- buttons ---
        btns = ctk.CTkFrame(parent, fg_color="transparent")
        btns.pack(pady=10)
        self.run_btn = ctk.CTkButton(btns, text="Run", width=120,
                                     command=self.run_clicked, state="disabled")
        self.run_btn.pack(side="left", padx=10)
        ctk.CTkButton(btns, text="Open Folder", width=120,
                      command=self.open_folder).pack(side="left", padx=10)

    # ---------- drag / browse ----------
    def browse(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if p:
            self.set_file(p)

    def on_drop(self, event):
        for f in self.drop_target.tk.splitlist(event.data):
            if f.lower().endswith(".xlsx"):
                self.set_file(f)
                break

    def set_file(self, path):
        self.file_path = path
        self.drop_info.configure(text=os.path.basename(path))
        self.run_btn.configure(state="normal")

    # ---------- run ----------
    def run_clicked(self):
        if not self.file_path:
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
        threading.Thread(target=self._worker,
                         args=(self.file_path, rows),
                         daemon=True).start()

    def _worker(self, src_path, rows):
        try:
            mawb  = splitter.get_mawb(src_path)
            df    = splitter.prepare_dataframe(src_path)
            parts = splitter.save_chunks(df, OUT_DIR, mawb, rows)
            messagebox.showinfo("Done", f"{parts} file(s) saved to:\n{OUT_DIR}")
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
        finally:
            self.run_btn.configure(state="normal")

    def open_folder(self):
        OUT_DIR.mkdir(exist_ok=True)
        os.startfile(OUT_DIR)


# ───────────────────────── main window ─────────────────────────
class MainApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        self.title(f"GA Broker Helper v{updater.__version__}")
        self.configure(bg="#1a1a1a")
        self.geometry("700x550")

        # grid: row-0 tabview, row-1 bottom controls
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # tabview
        self.tabview = ctk.CTkTabview(self, width=640, height=420)
        self.tabview.grid(row=0, column=0, columnspan=2,
                          padx=20, pady=(20, 10), sticky="nsew")

        ExcelSplitterTab(self.tabview.add("Excel Splitter"))
        RejectCodeSorterTab(self.tabview.add("Reject Code Sorter"))
        PGAReferenceTab(self.tabview.add("PGA Reference"))

        # ----- bottom-left banner -----
        if BANNER.exists():
            img = Image.open(BANNER)
            small = ctk.CTkImage(light_image=img, dark_image=img,
                                 size=(img.width//4, img.height//4))
            ctk.CTkLabel(self, image=small, text=""
                         ).grid(row=1, column=0, sticky="w",
                                padx=(20, 0), pady=(0, 20))
        else:
            ctk.CTkLabel(self, text=""
                         ).grid(row=1, column=0, sticky="w",
                                padx=(20, 0), pady=(0, 20))

        # ----- bottom-right update button -----
        ctk.CTkButton(self, text="Check for Update",
                      command=self._on_check_update, width=200
                      ).grid(row=1, column=1, sticky="e",
                             padx=(0, 20), pady=(0, 20))

    def _on_check_update(self):
        def worker():
            status = updater.check_and_update()
            if status == "latest":
                messagebox.showinfo("Updater", "You’re on the latest version.")
            elif status.startswith("error:"):
                messagebox.showerror("Updater", f"Update failed:\n{status[6:]}")
        threading.Thread(target=worker, daemon=True).start()


# ─────────────────────────── run ──────────────────────────────
if __name__ == "__main__":
    OUT_DIR.mkdir(exist_ok=True)
    MainApp().mainloop()
