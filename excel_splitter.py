# excel_splitter.py
"""
Excel Splitting Backend
------------------------
• Pure logic only
• No GUI code
• Expects a valid .xlsx path and splits by 998 rows
"""

import os, sys
import re
import pandas as pd
from openpyxl import load_workbook

APP_DIR = os.path.dirname(
    sys.executable if getattr(sys, "frozen", False) else __file__
    )
HEADER_PATH = os.path.join(APP_DIR, "Resources", "ExcelSplitter", "Header Sample.xlsx")
ROWS_PER_FILE = 499

HEADERS = [
    "Invoice_No", "Part", "Commercial_Description", "Country_of_Origin", "Country_of_Export",
    "Tariff_Number", "Quantity", "Quantity_UOM", "Unit_Price", "Total_Line_Value",
    "Net_Weight_KG", "Gross_Weight_KG", "Manufacturer_Name", "Manufacturer_Address_1",
    "Manufacturer_Address_2", "Manufacturer_City", "Manufacturer_State", "Manufacturer_Zip",
    "Manufacturer_Country", "MID_Code", "Buyer_Name", "Buyer_Address_1", "Buyer_Address_2",
    "Buyer_City", "Buyer_State", "Buyer_Zip", "Buyer_Country", "Buyer_ID_Number",
    "Consignee_Name", "Consignee_Address_1", "Consignee_Address_2", "Consignee_City",
    "Consignee_State", "Consignee_Zip", "Consignee_Country", "Consignee_ID_Number",
    "SICountry", "SP1", "SP2", "Zone_Status", "Privileged_Filing_Date", "Line_Piece_Count",
    "ADD_Case_Number", "CVD_Case_Number", "AD_Non_Reimbursement_Statement",
    "AD-CVD_Certification_Designation",
]

MAPPING = {
    "B": "F", "D": "G", "F": "J", "E": "L",
    "G": "M", "H": "N", "I": "P", "K": "R", "M": "T"
}

CONSTANTS = {
    "D": "CN", "E": "CN", "S": "CN", "H": "PCS"
}


def xl_idx(col: str) -> int:
    idx = 0
    for ch in col.upper():
        idx = idx * 26 + (ord(ch) - 64)
    return idx - 1


def get_mawb(path: str) -> str:
    wb = load_workbook(path, read_only=True, data_only=True)
    mawb = str(wb.active["U9"].value).strip()
    wb.close()
    return mawb


def prepare_dataframe(path: str) -> pd.DataFrame:
    raw = pd.read_excel(path, skiprows=9, engine="openpyxl")
    mapped = {
        HEADERS[xl_idx(tgt)]: raw.iloc[:, xl_idx(src)]
        for src, tgt in MAPPING.items()
    }
    tmp = pd.DataFrame(mapped).dropna(how="all").reset_index(drop=True)
    df = pd.DataFrame(index=tmp.index, columns=HEADERS)
    df.update(tmp)
    for tgt, val in CONSTANTS.items():
        df[HEADERS[xl_idx(tgt)]] = val

    mid_c, cntry_c = HEADERS[xl_idx("T")], HEADERS[xl_idx("S")]

    sg_mask = df[mid_c].astype(str).str.strip().str.upper().str.startswith("SG", na=False)
    df.loc[sg_mask, cntry_c] = "SG"

    hk_mask = df[mid_c].astype(str).str.strip().str.upper().str.startswith("HK", na=False)
    df.loc[hk_mask, cntry_c] = "HK"



    def _pad_zip(x):
        if pd.isna(x): return ""
        s = str(x).strip()
        m = re.fullmatch(r'(\d+)(?:\.0*)?', s)
        return f"{int(m.group(1)):06d}" if m else s

    for zip_col in ("Manufacturer_Zip", "Buyer_Zip"):
        if zip_col in df.columns:
            df[zip_col] = df[zip_col].apply(_pad_zip)

    return df


def save_chunks(df: pd.DataFrame, out_dir: str, mawb: str, rows: int = ROWS_PER_FILE) -> int:
    os.makedirs(out_dir, exist_ok=True)

    # ── choose suffix pattern based on rows-per-file ──
    import string
    if rows < 600:
        part_list = [f"{ltr}{i+1}" for ltr in string.ascii_uppercase for i in range(2)]  # A1,A2,B1,B2,...
    else:
        part_list = list(string.ascii_uppercase)  # A,B,C,...

    # capacity check
    if len(df) > rows * len(part_list):
        raise ValueError("Too many rows for available file parts.")

    # (rest of function is identical)

    if len(df) > rows * len(part_list):
        raise ValueError("Too many rows for available file parts.")

    template_headers = [
        cell.value for cell in load_workbook(HEADER_PATH, read_only=True).active[1] if cell.value
    ]

    part = 0
    for start in range(0, len(df), rows):
        chunk = df.iloc[start:start + rows].copy()
        suffix = part_list[part]
        invoice = f"{mawb}-{suffix}"
        chunk["Invoice_No"] = invoice
        chunk = chunk.reindex(columns=template_headers)
        path = os.path.join(out_dir, f"{invoice}.xlsx")
        with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
            chunk.to_excel(writer, index=False, header=False, startrow=1)
            worksheet = writer.sheets["Sheet1"]
            fmt = writer.book.add_format({'bold': True, 'font_name': 'Times New Roman', 'font_size': 12, 'align': 'center'})
            for i, h in enumerate(template_headers):
                worksheet.write(0, i, h, fmt)
            worksheet.set_column(0, 0, max(len(invoice), len("Invoice_No")) + 2)
            worksheet.set_column(5, 5, 20)
        part += 1
    return part
