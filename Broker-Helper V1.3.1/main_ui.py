import os, sys, threading
import customtkinter as ctk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
from PIL import Image
import excel_splitter as splitter
from reject_code_sorter import RejectCodeSorterTab
from pga_reference import PGAReferenceTab


APP_DIR = os.path.dirname(sys.executable if getattr(sys, "frozen", False) else __file__)
LOGO_DIR = os.path.join(APP_DIR, "Resources", "Logo")
OUT_DIR = os.path.join(APP_DIR, "splitted_excels")

BANNER = os.path.join(LOGO_DIR, "company_banner.png")
ICON   = os.path.join(LOGO_DIR, "company_logo.ico")

# 🟢 Force light theme (fixes mixed look)
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class ExcelSplitterTab:
    def __init__(self, parent):
        self.file_path = ""

        ctk.CTkLabel(parent, text="Drag & Drop Excel File Here or Use Browse",
                     font=("Arial", 14)).pack(pady=(20, 10))

        # ------ drop zone ------
        self.drop_target = ctk.CTkFrame(parent, height=60, width=400,
                                        fg_color="#808080", corner_radius=10)
        self.drop_target.pack(pady=5)
        self.drop_target.pack_propagate(False)

        self.drop_info = ctk.CTkLabel(self.drop_target, text="No file selected",
                                      font=("Arial", 12), text_color="#000000")
        self.drop_info.pack(expand=True)

        self.drop_target.drop_target_register(DND_FILES)
        self.drop_target.dnd_bind("<<Drop>>", self.on_drop)

        ctk.CTkButton(parent, text="Browse File", command=self.browse).pack(pady=5)

        # ------ rows-per-file row ------
        row_frame = ctk.CTkFrame(parent, fg_color="transparent")
        row_frame.pack(pady=4)
        ctk.CTkLabel(row_frame, text="Rows per file:").pack(side="left", padx=(0, 5))
        self.rows_var = ctk.StringVar(value="599")          # default 599
        ctk.CTkEntry(row_frame, width=80, textvariable=self.rows_var).pack(side="left")

        # ------ buttons ------
        btns = ctk.CTkFrame(parent, fg_color="transparent")
        btns.pack(pady=10)
        self.run_btn = ctk.CTkButton(btns, text="Run", width=120,
                                     command=self.run_clicked, state="disabled")
        self.run_btn.pack(side="left", padx=10)
        ctk.CTkButton(btns, text="Open Folder", width=120,
                      command=self.open_folder).pack(side="left", padx=10)

    # -------- drag / browse helpers --------
    def browse(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if p: self.set_file(p)

    def on_drop(self, event):
        for f in self.drop_target.tk.splitlist(event.data):
            if f.lower().endswith(".xlsx"):
                self.set_file(f); break

    def set_file(self, path):
        self.file_path = path
        self.drop_info.configure(text=os.path.basename(path))
        self.run_btn.configure(state="normal")

    # -------- worker thread --------
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
        threading.Thread(target=self._worker, args=(self.file_path, rows), daemon=True).start()

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
        os.makedirs(OUT_DIR, exist_ok=True)
        os.startfile(OUT_DIR)


class MainApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        self.configure(bg="#1a1a1a")  # <- dark background fix

        self.title("GA Broker • Toolkit")
        if os.path.exists(ICON):
            try:
                self.iconbitmap(ICON)
            except Exception:
                pass

        if os.path.exists(BANNER):
            banner_img = Image.open(BANNER)
            bw, bh = banner_img.size
            ctk.CTkLabel(
                self,
                image=ctk.CTkImage(light_image=banner_img, dark_image=banner_img, size=(bw, bh)),
                text=""
            ).pack(pady=(10, 20))
            self.geometry(f"{max(bw + 40, 640)}x530")
        else:
            self.geometry("700x460")

        self.resizable(False, False)

        tabview = ctk.CTkTabview(self, width=640, height=370)
        tabview.pack(padx=20, pady=10)

        ExcelSplitterTab(tabview.add("Excel Splitter"))
        RejectCodeSorterTab(tabview.add("Reject Code Sorter"))
        PGAReferenceTab(tabview.add("PGA Reference"))


if __name__ == "__main__":
    os.makedirs(OUT_DIR, exist_ok=True)
    MainApp().mainloop()
