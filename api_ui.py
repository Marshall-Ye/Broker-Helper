"""
api_ui.py ── Tkinter UI + network layer
Drag-and-drop Excel → build payload via json_content.build_payload
→ POST to Acelynk.  Detailed server reply is saved to last_reply.json.
"""

from __future__ import annotations
import json, logging, os, threading
from pathlib import Path

import pandas as pd
import requests
import customtkinter as ctk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES

from json_content import build_entry   # ← all JSON tweaks live there

# ───────── static config ─────────
API = "https://cert-api.acelynk.com"
TOKEN_ENDPOINT = f"{API}/token"
POST_ENDPOINT  = f"{API}/api/Invoices/CreateEntrySummary"

LAST_PAYLOAD = Path(__file__).with_name("last_payload.json")
LAST_REPLY   = Path(__file__).with_name("last_reply.json")

USERNAME = "cert_acelynk_D9T"
PASSWORD = "Cg4fz9Sk0jNuL6"

logging.basicConfig(level=logging.DEBUG,
                    format="%(asctime)s  %(levelname)s  %(message)s")


# ───────── Tk Tab ─────────
class JsonConverterTab:
    def __init__(self, parent, output_dir: Path):
        self.file_path: str = ""

        ctk.CTkLabel(parent,
                     text="Drag split Excel → create entry header",
                     font=("Arial", 14)).pack(pady=(18, 10))

        drop = ctk.CTkFrame(parent, width=420, height=60,
                            fg_color="#808080", corner_radius=10)
        drop.pack(pady=5)
        drop.pack_propagate(False)

        self.drop_lbl = ctk.CTkLabel(drop, text="No file selected",
                                     text_color="#000000")
        self.drop_lbl.pack(expand=True)

        drop.drop_target_register(DND_FILES)
        drop.dnd_bind("<<Drop>>", self._on_drop)

        ctk.CTkButton(parent, text="Browse file",
                      command=self._browse).pack(pady=5)

        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(pady=10)
        self.run_btn = ctk.CTkButton(
            row, text="Upload entry",
            command=self._on_run, state="disabled", width=160
        )
        self.run_btn.pack(side="left", padx=10)

        # quick-open buttons for payload & reply
        ctk.CTkButton(row, text="Open payload",
                      command=lambda: os.startfile(LAST_PAYLOAD)
                      ).pack(side="left", padx=6)
        ctk.CTkButton(row, text="Open reply",
                      command=lambda: os.startfile(LAST_REPLY)
                      ).pack(side="left", padx=6)

        self.log = ctk.CTkTextbox(parent, width=420, height=120, state="disabled")
        self.log.pack(pady=(0, 12))

    # ---------- drag / browse ----------
    def _browse(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if p:
            self._set_file(p)

    def _on_drop(self, ev):
        for f in ev.data.split():
            if f.lower().endswith(".xlsx"):
                self._set_file(f)
                break

    def _set_file(self, p: str):
        self.file_path = p
        self.drop_lbl.configure(text=os.path.basename(p))
        self.run_btn.configure(state="normal")

    # ---------- run ----------
    def _on_run(self):
        if not self.file_path:
            messagebox.showerror("Missing file", "Pick an Excel file"); return
        self.run_btn.configure(state="disabled"); self._log_clear()
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        try:
            self._gui(lambda: self._log("reading Excel…"))
            df = pd.read_excel(self.file_path, engine="openpyxl")
            first_row = df.iloc[0]

            self._gui(lambda: self._log("building JSON…"))
            payload = build_entry(first_row)
            LAST_PAYLOAD.write_text(json.dumps(payload, indent=2, ensure_ascii=False))

            token = self._token()
            self._gui(lambda: self._log("posting to Acelynk…"))
            reply = self._post_and_log(payload, token)

            # Decide success
            ok = False
            if isinstance(reply, dict):
                # old style { response: [ {StatusCode: 200, ...} ] }
                inner = reply.get("response", [{}])[0]
                ok = inner.get("StatusCode") == 200
            # fallback: HTTP 200 with no body error
            if ok:
                self._gui(lambda: self._log("✅ entry accepted"))
                messagebox.showinfo("Done", "Entry created - check Unassigned.")
            else:
                raise RuntimeError("See last_reply.json for details.")

        except Exception as exc:
            self._gui(lambda: self._log(f"❌ {exc}"))
            messagebox.showerror("Error", str(exc))
        finally:
            self._gui(lambda: self.run_btn.configure(state="normal"))

    # ---------- network ----------
    def _token(self) -> str:
        r = requests.post(
            TOKEN_ENDPOINT,
            data={"userName": USERNAME,
                  "password": PASSWORD,
                  "grant_type": "password"},
            timeout=15
        )
        r.raise_for_status()
        return r.json()["access_token"]

    def _post_and_log(self, payload: dict, token: str):
        r = requests.post(
            POST_ENDPOINT, json=payload,
            headers={"Authorization": f"Bearer {token}"},
            timeout=30
        )

        # Always write raw body for debugging (even on 500)
        try:
            body = r.json()
        except ValueError:
            body = r.text

        LAST_REPLY.write_text(json.dumps(body, indent=2, ensure_ascii=False)
                              if isinstance(body, (dict, list)) else str(body))

        cid = r.headers.get("X-Correlation-ID", "n/a")
        self._gui(lambda: self._log(f"HTTP {r.status_code}  CorrelationID: {cid}"))
        return body

    # ---------- tiny ui helpers ----------
    def _log(self, m: str):
        self.log.configure(state="normal")
        self.log.insert("end", m + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _log_clear(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _gui(self, fn):
        self.run_btn.after(0, fn)
