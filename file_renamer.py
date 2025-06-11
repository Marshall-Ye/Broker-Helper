# file_renamer.py
"""
GA Broker Helper – File Renamer Tab
• Rename packing list into 'Packing List Renamed'
• Parse user-pasted invoice ↔ entry mapping
• Rename 3461 PDFs → '3461 Renamed/' (from original folder)
• Rename 7501 PDFs → '7501 Renamed/' (from original folder)
"""
from __future__ import annotations

import os, platform, re, shutil, subprocess, threading, datetime
from pathlib import Path
from typing import Optional

import customtkinter as ctk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES

# ──────────────── helpers ─────────────────────────────
def _open_path(path: Path) -> None:
    if platform.system() == "Windows":
        os.startfile(path)  # type: ignore[attr-defined]
    elif platform.system() == "Darwin":
        subprocess.call(["open", path])
    else:
        subprocess.call(["xdg-open", path])

_MAWB_RE = re.compile(r"\b(\d{3}-\d{8})\b")
_ENTRY_PATTERN = re.compile(r"(\d{8})")


def _rename_packing_list(src_dir: Path) -> tuple[Path, str, str]:
    excels = list(src_dir.glob("*.xls*"))
    if len(excels) != 1:
        raise FileNotFoundError("Expected 1 Excel packing list in folder.")

    pl_path = excels[0]
    mawb_match = _MAWB_RE.search(pl_path.stem)
    if not mawb_match:
        raise ValueError("MAWB not found in packing list filename.")
    mawb = mawb_match.group(1)
    date_str = datetime.date.today().strftime("%Y-%m-%d")

    out_root = src_dir.parent / f"{mawb} renamed"
    (out_root / "Packing List Renamed").mkdir(parents=True, exist_ok=True)
    (out_root / "3461 Renamed").mkdir(exist_ok=True)
    (out_root / "7501 Renamed").mkdir(exist_ok=True)

    new_name = f"GA_PL_{mawb}-1_{date_str}{pl_path.suffix}"
    shutil.copy2(pl_path, out_root / "Packing List Renamed" / new_name)

    return out_root, mawb, date_str


def _parse_mapping(text: str, mawb: str) -> dict[str, str]:
    pattern = re.compile(
        rf"{re.escape(mawb)}-([A-Z])\s+([0-9]{{8}})"
    )
    mapping: dict[str, str] = {}
    for line in text.splitlines():
        m = pattern.search(line)
        if m:
            letter = m.group(1)
            entry = m.group(2).lstrip("0")
            mapping[entry] = letter

    if not mapping:
        raise ValueError("No MAWB-letter / entry-number pairs found.")
    return mapping


def _rename_3461_pdfs(original_root: Path, output_root: Path, mawb: str, date_str: str, mapping: dict[str, str]) -> int:
    src_dir = original_root / "3461"
    dest_dir = output_root / "3461 Renamed"
    count = 0

    if not src_dir.exists():
        raise FileNotFoundError(f"Missing folder:\n{src_dir}")

    for file in src_dir.glob("*.pdf"):
        match = _ENTRY_PATTERN.search(file.stem)
        if not match:
            continue
        entry = match.group(1).lstrip("0")
        invoice = mapping.get(entry)
        if not invoice:
            continue
        new_name = f"GA_CF3461_{mawb}-{invoice}_{date_str}.pdf"
        shutil.copy2(file, dest_dir / new_name)
        count += 1

    return count


def _rename_7501_pdfs(original_root: Path, output_root: Path, mawb: str, date_str: str, mapping: dict[str, str]) -> int:
    src_dir = original_root / "7501"
    dest_dir = output_root / "7501 Renamed"
    count = 0

    if not src_dir.exists():
        raise FileNotFoundError(f"Missing folder:\n{src_dir}")

    for file in src_dir.glob("*.pdf"):
        match = _ENTRY_PATTERN.search(file.stem)
        if not match:
            continue
        entry = match.group(1).lstrip("0")
        invoice = mapping.get(entry)
        if not invoice:
            continue
        new_name = f"GA_CF7501_{mawb}-{invoice}_{date_str}.pdf"
        shutil.copy2(file, dest_dir / new_name)
        count += 1

    return count

# ──────────────── GUI tab ─────────────────────────────
class FileRenamerTab:
    def __init__(self, parent: ctk.CTkFrame):
        self._parent = parent
        self._folder: Optional[Path] = None
        self._out_dir: Optional[Path] = None
        self._mapping: dict[str, str] = {}

        ctk.CTkLabel(
            parent,
            text="Drag & Drop a shipment folder here (or click Browse)",
            font=("Arial", 14), anchor="w", wraplength=600,
        ).pack(padx=20, pady=(20, 10))

        self._drop = ctk.CTkFrame(parent, width=420, height=70,
                                  fg_color="#808080", corner_radius=10)
        self._drop.pack(pady=6)
        self._drop.pack_propagate(False)

        self._drop_msg = ctk.CTkLabel(
            self._drop, text="No folder selected",
            font=("Arial", 12), text_color="#000000"
        )
        self._drop_msg.pack(expand=True)

        self._drop.drop_target_register(DND_FILES)
        self._drop.dnd_bind("<<Drop>>", self._on_drop)

        ctk.CTkButton(parent, text="Browse Folder", command=self._browse).pack(pady=6)

        ctk.CTkLabel(
            parent,
            text="Paste Invoice ↔ Entry mapping below, then click Rename:",
            font=("Arial", 12), anchor="w"
        ).pack(padx=20, pady=(12, 4), fill="x")

        self._map_text = ctk.CTkTextbox(parent, width=600, height=120, wrap="none")
        self._map_text.pack(padx=20, pady=(0, 8), fill="both", expand=False)

        btn_row = ctk.CTkFrame(parent, fg_color="transparent")
        btn_row.pack(pady=(10, 4))

        self._run_btn = ctk.CTkButton(btn_row, text="Rename", width=140,
                                      command=self._on_run, state="disabled")
        self._run_btn.pack(side="left", padx=10)

        self._open_btn = ctk.CTkButton(btn_row, text="Open Folder", width=140,
                                       command=self._open_folder, state="disabled")
        self._open_btn.pack(side="left", padx=10)

    def _browse(self):
        p = filedialog.askdirectory(title="Select shipment folder")
        if p:
            self._set_folder(Path(p))

    def _on_drop(self, event):
        first = self._drop.tk.splitlist(event.data)[0]
        if os.path.isdir(first):
            self._set_folder(Path(first))

    def _set_folder(self, path: Path):
        self._folder = path
        self._drop_msg.configure(text=path.name)
        self._run_btn.configure(state="normal")
        self._open_btn.configure(state="disabled")
        self._out_dir = None
        self._mapping.clear()

    def _on_run(self):
        if not self._folder:
            messagebox.showerror("No folder selected", "Please select a shipment folder first.")
            return

        mapping_text = self._map_text.get("1.0", "end")
        self._run_btn.configure(state="disabled")

        def worker(txt: str):
            try:
                out_root, mawb, date_str = _rename_packing_list(self._folder)
                self._mapping = _parse_mapping(txt, mawb)

                print("\n▶ Invoice-letter ↔ Entry-number pairs")
                for entry, letter in self._mapping.items():
                    print(f"  {entry}  ⇄  {letter}")
                print(f"Total pairs parsed: {len(self._mapping)}\n")

                renamed_3461 = _rename_3461_pdfs(self._folder, out_root, mawb, date_str, self._mapping)
                renamed_7501 = _rename_7501_pdfs(self._folder, out_root, mawb, date_str, self._mapping)

                print(f"✔ Renamed {renamed_3461} PDF(s) under '3461 Renamed/'")
                print(f"✔ Renamed {renamed_7501} PDF(s) under '7501 Renamed/'")

            except Exception as exc:
                self._async(lambda exc=exc: messagebox.showerror("Renamer", str(exc)))
            else:
                self._out_dir = out_root
                self._async(lambda: [
                    messagebox.showinfo("Renamer",
                        f"Packing list renamed.\n"
                        f"Parsed {len(self._mapping)} mapping pair(s).\n"
                        f"Renamed {renamed_3461} 3461 PDF(s).\n"
                        f"Renamed {renamed_7501} 7501 PDF(s).\n"
                        f"Output folder:\n{out_root}"),
                    self._open_btn.configure(state="normal"),
                ])
            finally:
                self._async(lambda: self._run_btn.configure(state="normal"))

        threading.Thread(target=worker, args=(mapping_text,), daemon=True).start()

    def _open_folder(self):
        if self._out_dir and self._out_dir.exists():
            _open_path(self._out_dir)
        else:
            messagebox.showwarning("Renamer", "No renamed-files folder available yet.")

    def _async(self, fn, ms: int = 0):
        self._parent.after(ms, fn)
