"""
reject_code_sorter.py  –  PDF “Reject Code” extractor  (v2)
-----------------------------------------------------------
• Drag-and-drop / Browse a PDF
• Groups 'Line# xxx yyy' entries by message-ID
• Creates /generated_txts/<pdf>.txt with:
      – side-notes on every ID
      – 465 placed just above 628, 628 last
• GUI tab class: RejectCodeSorterTab (dark style, same colors as before)
"""

import os, re, sys, threading
import fitz                              # PyMuPDF
import customtkinter as ctk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES
from collections import defaultdict

# ── paths ────────────────────────────────────────────────────
APP_DIR     = os.path.dirname(sys.executable if getattr(sys, "frozen", False) else __file__)
TXT_OUT_DIR = os.path.join(APP_DIR, "generated_txts")
os.makedirs(TXT_OUT_DIR, exist_ok=True)

# ── side-notes per message-ID ───────────────────────────────
SIDE_NOTE = {
    "628": "ignore",
    "465": "ignore",
    "523": "fix MID",
    "771": "add tariff: 9903.01.63",
    "794": "add CN",
    "687": "delete the line",
    "483": "calculate MID",
    "775": "delete the line",
    "613": "delete the line",
    "773": "change country to SG",
}

# ── core logic ───────────────────────────────────────────────
def read_pdf_to_txt(pdf_path: str) -> str:
    """Parse *pdf_path* and write an ordered .txt with side-notes."""
    doc = fitz.open(pdf_path)
    record_list = []

    # collect consecutive 'Line#' blocks
    for page in doc:
        matches = re.findall(r'(Line# \d+\s+\d+)', page.get_text())
        for m in matches:
            ln_no = int(re.findall(r'Line# (\d+)\s+\d+', m)[0])
            if not record_list or ln_no in (record_list[-1][0], record_list[-1][0] + 1):
                record_list.append([ln_no, m])

    # preserve first-seen order of IDs
    ordered_ids, seen = [], set()
    for _, raw in record_list:
        mid = re.findall(r'Line# \d+\s+(\d+)', raw)[0]
        if mid not in seen:
            ordered_ids.append(mid)
            seen.add(mid)

    # move 465 just before 628, 628 last
    ordered_ids = [i for i in ordered_ids if i not in ("465", "628")]
    if "465" in seen:
        ordered_ids.append("465")
    if "628" in seen:
        ordered_ids.append("628")

    # group line numbers by ID
    groups = defaultdict(list)
    for _, raw in record_list:
        _, ln, mid = raw.strip().split()
        groups[mid].append(f"Line# {ln}")

    # write file
    out_path = os.path.join(
        TXT_OUT_DIR, os.path.splitext(os.path.basename(pdf_path))[0] + ".txt"
    )
    with open(out_path, "w", encoding="utf-8") as f:
        for mid in ordered_ids:
            note = SIDE_NOTE.get(mid, "")
            f.write(f"\n{mid} {note}\n".rstrip() + "\n")   # header with side-note
            for ln in groups[mid]:
                f.write(f"{ln}\n")

    return out_path

# ── GUI tab ──────────────────────────────────────────────────
class RejectCodeSorterTab:
    def __init__(self, parent):
        self.pdf_path = ""

        ctk.CTkLabel(parent, text="Drag & Drop PDF Here or Use Browse",
                     font=("Arial", 14)).pack(pady=(20, 10))

        # drop zone (colors unchanged)
        self.drop_target = ctk.CTkFrame(parent, height=60, width=400,
                                        fg_color="#808080", corner_radius=10)
        self.drop_target.pack(pady=5)
        self.drop_target.pack_propagate(False)

        self.drop_info = ctk.CTkLabel(self.drop_target, text="No file selected",
                                      font=("Arial", 12), text_color="#000000")
        self.drop_info.pack(expand=True)

        self.drop_target.drop_target_register(DND_FILES)
        self.drop_target.dnd_bind("<<Drop>>", self._on_drop)

        ctk.CTkButton(parent, text="Browse File", command=self._browse).pack(pady=5)

        btns = ctk.CTkFrame(parent, fg_color="transparent")
        btns.pack(pady=10)
        self.run_btn = ctk.CTkButton(btns, text="Run", width=120,
                                     state="disabled", command=self._run_clicked)
        self.run_btn.pack(side="left", padx=10)
        ctk.CTkButton(btns, text="Open Folder", width=120,
                      command=lambda: os.startfile(TXT_OUT_DIR)).pack(side="left", padx=10)

    # ---------- helper callbacks ----------
    def _set_file(self, path):
        self.pdf_path = path
        self.drop_info.configure(text=os.path.basename(path))
        self.run_btn.configure(state="normal")

    def _browse(self):
        p = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if p: self._set_file(p)

    def _on_drop(self, event):
        for f in self.drop_target.tk.splitlist(event.data):
            if f.lower().endswith(".pdf"):
                self._set_file(f)
                break

    def _run_clicked(self):
        if not self.pdf_path:
            messagebox.showerror("No file selected", "Please pick a PDF.")
            return
        self.run_btn.configure(state="disabled")
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        try:
            out = read_pdf_to_txt(self.pdf_path)
            # instead of a messagebox, just open the .txt:
            os.startfile(out)
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.run_btn.configure(state="normal")
