# json_converter.py  ── GA Broker Helper
# ---------------------------------------------------------------
# Drag-and-drop split Excel → build minimal "RD" payload
# Upload to Acelynk, logging Correlation ID on any error
# Keeps last_payload.json for support
# ---------------------------------------------------------------

from __future__ import annotations
import json, logging, math, os, threading
from pathlib import Path
from typing import Any

import numpy as np, pandas as pd, requests, customtkinter as ctk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES

# ───────── API endpoints ─────────
API_BASE          = "https://cert-api.acelynk.com"
TOKEN_ENDPOINT    = f"{API_BASE}/token"
CREATE_ENDPOINT   = f"{API_BASE}/api/Invoices/CreateEntrySummary"
LAST_PAYLOAD      = Path(__file__).with_name("last_payload.json")

USERNAME = "cert_acelynk_D9T"
PASSWORD = "Cg4fz9Sk0jNuL6"

logging.basicConfig(level=logging.DEBUG,
                    format="%(asctime)s  %(levelname)s  %(message)s")

# ───────── NaN helpers ─────────
def _is_nan(val: Any) -> bool:
    return (
        val is None or
        (isinstance(val, float) and math.isnan(val)) or
        (isinstance(val, np.floating) and np.isnan(val)) or
        (isinstance(val, str) and val.strip().lower() in {"", "nan", "nat", "none"})
    )

def _strip_nans(obj: Any) -> Any:
    if isinstance(obj, dict):
        clean = {k: _strip_nans(v) for k, v in obj.items() if not _is_nan(v)}
        return {k: v for k, v in clean.items() if v not in ({}, [])}
    if isinstance(obj, list):
        return [_strip_nans(v) for v in obj if not _is_nan(v)]
    return "" if _is_nan(obj) else obj

# ───────── Tk Tab class ─────────
class JsonConverterTab:
    def __init__(self, parent, output_dir: Path):
        self.output_dir = Path(output_dir); self.output_dir.mkdir(exist_ok=True)
        self.file_path: str = ""

        ctk.CTkLabel(parent, text="Drag Excel → upload line items to Acelynk",
                     font=("Arial", 14)).pack(pady=(20, 10))

        self.drop_frame = ctk.CTkFrame(parent, width=420, height=60,
                                       fg_color="#808080", corner_radius=10)
        self.drop_frame.pack(pady=5); self.drop_frame.pack_propagate(False)
        self.drop_label = ctk.CTkLabel(self.drop_frame, text="No file selected",
                                       text_color="#000000")
        self.drop_label.pack(expand=True)
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind("<<Drop>>", self._on_drop)
        ctk.CTkButton(parent, text="Browse file", command=self._browse).pack(pady=5)

        entry_row = ctk.CTkFrame(parent, fg_color="transparent"); entry_row.pack(pady=(8,0))
        ctk.CTkLabel(entry_row, text="Entry #").pack(side="left", padx=(0,6))
        self.entry_var = ctk.StringVar(value="00000012")     # pre-fill
        ctk.CTkEntry(entry_row, width=130, textvariable=self.entry_var).pack(side="left")

        btn_row = ctk.CTkFrame(parent, fg_color="transparent"); btn_row.pack(pady=10)
        self.run_btn = ctk.CTkButton(btn_row, text="Upload line items",
                                     command=self._on_run, state="disabled", width=150)
        self.run_btn.pack(side="left", padx=10)
        ctk.CTkButton(btn_row, text="Open payload file",
                      command=lambda: os.startfile(LAST_PAYLOAD)).pack(side="left", padx=10)

        self.log_box = ctk.CTkTextbox(parent, width=420, height=120, state="disabled")
        self.log_box.pack(pady=(0,12))

    # ---------- UI helpers ----------
    def _browse(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if p: self._set_file(p)

    def _on_drop(self, event):
        for f in self.drop_frame.tk.splitlist(event.data):
            if f.lower().endswith(".xlsx"):
                self._set_file(f); break

    def _set_file(self, path: str):
        self.file_path = path
        self.drop_label.configure(text=os.path.basename(path))
        self.run_btn.configure(state="normal")

    # ---------- action ----------
    def _on_run(self):
        if not self.file_path:
            messagebox.showerror("Missing file", "Pick an Excel file first."); return
        if not self.entry_var.get().strip():
            messagebox.showerror("Missing entry #", "Type the entry number."); return
        self.run_btn.configure(state="disabled"); self._log_clear()
        threading.Thread(target=self._worker, daemon=True).start()

    # ---------- worker ----------
    def _worker(self):
        try:
            self._gui(lambda: self._log("reading Excel…"))
            df = pd.read_excel(self.file_path, engine="openpyxl")

            self._gui(lambda: self._log("building JSON…"))
            payload = _strip_nans(self._build_json(df))
            LAST_PAYLOAD.write_text(json.dumps(payload, indent=2, ensure_ascii=False))
            self._gui(lambda: self._log(f"payload saved → {LAST_PAYLOAD.name}"))

            token = self._get_token()

            self._gui(lambda: self._log("uploading to Acelynk…"))
            self._post_and_check(CREATE_ENDPOINT, payload, token)

            self._gui(lambda: self._log("✅ upload OK"))
            self._gui(lambda: messagebox.showinfo("Done", "Upload complete."))
        except Exception as exc:
            err = str(exc)
            self._gui(lambda m=err: self._log(f"❌ {m}"))
            self._gui(lambda m=err: messagebox.showerror("Error", m))
        finally:
            self._gui(lambda: self.run_btn.configure(state="normal"))

    # ---------- HTTP helper ----------
    def _post_and_check(self, url: str, payload: dict, token: str):
        resp = requests.post(url, json=payload,
                             headers={"Authorization": f"Bearer {token}"}, timeout=30)
        if resp.status_code >= 400:
            corr = (resp.headers.get("X-Correlation-ID")
                    or resp.headers.get("Request-CorrelationID", "n/a"))
            err = (f"Status: {resp.status_code}\n"
                   f"CorrelationID: {corr}\n"
                   f"Body: {resp.text}")
            raise RuntimeError(err)

    # ---------- tariff helper ----------
    @staticmethod
    def _clean_tariff(val: Any) -> str:
        digits = "".join(ch for ch in str(val) if ch.isdigit())
        return digits.zfill(10)[:10]

    # ---------- JSON builder ----------
    def _build_json(self, df: pd.DataFrame) -> dict:
        """
        Build an 'RD' payload that matches Acelynk's canonical casing / nesting:
          • Id  (001, 002 …)
          • Inv_number   (PascalCase)
          • Manufacturer  as an object
          • Quantity / Price / Tariff objects with correct inner keys
        """
        invoice_no = str(df.iloc[0]["Invoice_No"]).strip()

        # ----- hard-coded importer block -----
        importer = {
            "name": "GOLDEN ARCUS INTL CO",
            "address": {
                "address_1": "5343 W IMPERIAL HWY NO. 700",
                "address_2": ""
            },
            "city": "LOS ANGELES",
            "region": "CA",
            "country": "US",
            "postal_code": "90045",
            # EIN 46-2750048 + suffix 00  →  46275004800  (no dashes)
            "account": "46275004800",
            "tax_id": "",
            "filer": ""
        }

        # ----- build Line-Item list -----
        items: list[dict[str, Any]] = []
        for idx, row in df.iterrows():
            items.append({
                "Id": f"{idx + 1:03d}",  # "001", "002", …
                "Inv_number": invoice_no,
                "Description": str(row["Commercial_Description"]).strip(),
                "Product_number": str(row["Part"]).strip(),

                "Country_of_origin": row["Country_of_Origin"],
                "Country_of_export": row["Country_of_Export"],

                # --- Trade Agreement omitted (add if needed) ---

                "Manufacturer": {
                    "Name": str(row["Manufacturer_Name"]).strip(),
                    "Address": {
                        "Address_1": str(row["Manufacturer_Address_1"]).strip()
                    },
                    "Mid_code": str(row["MID_Code"]).strip()
                },

                "Tariff": [{
                    "Number": self._clean_tariff(row["Tariff_Number"]),
                    "Reporting_quantity_1": {
                        "Uom": row["Quantity_UOM"],
                        "Text": f"{float(row['Quantity']):.0f}"
                    }
                }],

                "Quantity": {
                    "Amount": f"{float(row['Quantity']):.0f}",
                    "Uom": row["Quantity_UOM"]
                },

                "Price": {
                    "Unit_price": f"{float(row['Unit_Price']):.2f}",
                    "Total_price": f"{float(row['Total_Line_Value']):.2f}",
                    "Name": "Total Value"  # required by schema
                }
            })

        # ----- assemble final payload -----
        return {
            "shipment": [{
                "header": {
                    "general": {"invoice_number": invoice_no},
                    "importer_of_record": importer,

                    # ---- NEW: terms & freight blocks ----
                    "terms": {
                        "terms_of_sale": "FOB",
                        "terms_location": ""
                    },
                    "freight": {
                        "freight_included_in_invoice": True,
                        "charges": "0.00",
                        "currency": "USD"
                    }
                },
                "line": {"item": items},
                "invoice_number": invoice_no,
                "user_action_code": "RD"  # no entry_number – match by invoice
            }]
        }

    # ---------- auth helper ----------
    def _get_token(self) -> str:
        resp = requests.post(TOKEN_ENDPOINT,
                             data={"userName": USERNAME,
                                   "password": PASSWORD,
                                   "grant_type": "password"},
                             timeout=15)
        resp.raise_for_status()
        return resp.json()["access_token"]

    # ---------- logging ----------
    def _log(self, msg: str):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", msg + "\n"); self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _log_clear(self):
        self.log_box.configure(state="normal"); self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    # ---------- Tk marshal ----------
    def _gui(self, func): self.run_btn.after(0, func)