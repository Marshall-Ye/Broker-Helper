# json_converter.py ─ GA Broker Helper
# ---------------------------------------------------------------
# Drag-and-drop split Excel → post FULL entry to Acelynk
# Identified by invoice_number (creates or replaces the entry)
# Saves last_payload.json and logs correlation IDs on errors
# ---------------------------------------------------------------

from __future__ import annotations
import json, logging, math, os, threading
from pathlib import Path
from typing import Any

import numpy as np, pandas as pd, requests, customtkinter as ctk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES

# ───────── API endpoints ─────────
API = "https://cert-api.acelynk.com"
TOKEN_ENDPOINT = f"{API}/token"
POST_ENDPOINT  = f"{API}/api/Invoices/CreateEntrySummary"
LAST_PAYLOAD   = Path(__file__).with_name("last_payload.json")

USERNAME = "cert_acelynk_D9T"
PASSWORD = "Cg4fz9Sk0jNuL6"

logging.basicConfig(level=logging.DEBUG,
                    format="%(asctime)s  %(levelname)s  %(message)s")

# ───────── helpers ─────────
def _is_nan(v: Any) -> bool:
    return (
        v is None or
        (isinstance(v, float) and math.isnan(v)) or
        (isinstance(v, np.floating) and np.isnan(v)) or
        (isinstance(v, str) and v.strip().lower() in {"", "nan", "nat", "none"})
    )

def _strip_nans(o: Any) -> Any:
    if isinstance(o, dict):
        cleaned = {k: _strip_nans(v) for k, v in o.items() if not _is_nan(v)}
        return {k: v for k, v in cleaned.items() if v != {}}  # keep [] !
    if isinstance(o, list):
        return [_strip_nans(v) for v in o if not _is_nan(v)]
    return "" if _is_nan(o) else o

def _clean_tariff(v: Any) -> str:
    digits = "".join(c for c in str(v) if c.isdigit())
    return digits.zfill(10)[:10]

# ───────── Tk tab ─────────
class JsonConverterTab:
    def __init__(self, parent, output_dir: Path):
        self.output_dir = Path(output_dir); self.output_dir.mkdir(exist_ok=True)
        self.file_path: str = ""

        ctk.CTkLabel(parent, text="Drag Excel → create / replace entry in Acelynk",
                     font=("Arial", 14)).pack(pady=(20,10))

        # drag-and-drop frame
        self.drop = ctk.CTkFrame(parent, width=420, height=60,
                                 fg_color="#808080", corner_radius=10)
        self.drop.pack(pady=5); self.drop.pack_propagate(False)
        self.drop_lbl = ctk.CTkLabel(self.drop, text="No file selected", text_color="#000000")
        self.drop_lbl.pack(expand=True)
        self.drop.drop_target_register(DND_FILES)
        self.drop.dnd_bind("<<Drop>>", self._on_drop)

        ctk.CTkButton(parent, text="Browse file", command=self._browse).pack(pady=5)

        # action button row
        row = ctk.CTkFrame(parent, fg_color="transparent"); row.pack(pady=10)
        self.run_btn = ctk.CTkButton(row, text="Upload entry",
                                     command=self._on_run, state="disabled", width=160)
        self.run_btn.pack(side="left", padx=10)
        ctk.CTkButton(row, text="Open payload file",
                      command=lambda: os.startfile(LAST_PAYLOAD)).pack(side="left", padx=10)

        self.log = ctk.CTkTextbox(parent, width=420, height=120, state="disabled")
        self.log.pack(pady=(0,12))

    # ---- UI helpers ----
    def _browse(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx")])
        if p: self._set_file(p)

    def _on_drop(self, ev):
        for f in self.drop.tk.splitlist(ev.data):
            if f.lower().endswith(".xlsx"):
                self._set_file(f); break

    def _set_file(self, path):
        self.file_path = path
        self.drop_lbl.configure(text=os.path.basename(path))
        self.run_btn.configure(state="normal")

    # ---- run ----
    def _on_run(self):
        if not self.file_path:
            messagebox.showerror("Missing file","Pick an Excel file first"); return
        self.run_btn.configure(state="disabled"); self._log_clear()
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        try:
            self._gui(lambda: self._log("reading Excel…"))
            df = pd.read_excel(self.file_path, engine="openpyxl")

            self._gui(lambda: self._log("building JSON…"))
            payload = _strip_nans(self._build_json(df))
            LAST_PAYLOAD.write_text(json.dumps(payload, indent=2, ensure_ascii=False))
            self._gui(lambda: self._log(f"payload saved → {LAST_PAYLOAD.name}"))

            token = self._get_token()
            self._gui(lambda: self._log("posting to Acelynk…"))
            self._post_and_check(payload, token)

            self._gui(lambda: self._log("✅ 200 OK – entry created/replaced"))
            messagebox.showinfo("Done","Upload complete (see Unassigned or Filer queue).")
        except Exception as e:
            err = str(e)
            self._gui(lambda: self._log(f"❌ {err}"))
            messagebox.showerror("Error", err)
        finally:
            self._gui(lambda: self.run_btn.configure(state="normal"))

    # ---------- JSON builder ----------
    def _build_json(self, df: pd.DataFrame) -> dict:
        invoice_no = str(df.iloc[0]["Invoice_No"]).strip()

        # 1️⃣  Importer with only Account + Filer  (per minimal spec)
        importer = {
            "Account": "46275004800",  # EIN without dashes
            "Filer": "D9T"
        }

        # 2️⃣  Header → General + Importer blocks
        header = {
            "Importer_of_record": importer,
            "General": {
                "Invoice_number": invoice_no,
                "BillLading": []  # keep as [] (Magaya example)
            }
        }

        # 3️⃣  Build super-lean line items -------------
        items: list[dict[str, Any]] = []
        for row in df.itertuples(index=False):
            items.append({
                "inv_number": invoice_no,  # lower-case per sample
                "Product_number": str(row.Part),
                "Country_of_origin": row.Country_of_Origin,
                "Country_of_export": row.Country_of_Export,
                "Tariff": [{
                    "Number": _clean_tariff(row.Tariff_Number)
                }]
                # (quantity / price / etc. omitted – add if CBP requires)
            })

        # 4️⃣  Assemble payload with Reference right above Line
        shipment = {
            "Reference": invoice_no,  # moved here
            "Header": header,
            "Line": {"Item": items}
        }

        return {"Shipment": [shipment]}

    # ---- networking helpers ----
    def _get_token(self) -> str:
        r = requests.post(TOKEN_ENDPOINT,
            data={"userName": USERNAME,
                  "password": PASSWORD,
                  "grant_type": "password"}, timeout=15)
        r.raise_for_status()
        return r.json()["access_token"]

    def _post_and_check(self, payload, token):
        r = requests.post(POST_ENDPOINT, json=payload,
                          headers={"Authorization": f"Bearer {token}"}, timeout=30)
        if r.status_code >= 400:
            cid = r.headers.get("X-Correlation-ID") or "n/a"
            raise RuntimeError(f"Status: {r.status_code}\nCorrelationID: {cid}\nBody: {r.text}")

    # ---- GUI log helpers ----
    def _log(self, msg): self.log.configure(state="normal"); self.log.insert("end", msg+"\n"); self.log.see("end"); self.log.configure(state="disabled")
    def _log_clear(self): self.log.configure(state="normal"); self.log.delete("1.0","end"); self.log.configure(state="disabled")
    def _gui(self, fn): self.run_btn.after(0, fn)
