# excel_splitter.py
"""
Excel Splitting Backend
------------------------
• Pure logic only
• No GUI code
• Expects a valid .xlsx path and splits by ROWS_PER_FILE rows
"""

import os
import re
import string
import pandas as pd
from openpyxl import load_workbook

APP_DIR     = os.path.dirname(__file__)
HEADER_PATH = os.path.join(APP_DIR, "Resources", "ExcelSplitter", "Header Sample.xlsx")
ROWS_PER_FILE = 499

HEADERS = [
    "Invoice_No","Part","Commercial_Description","Country_of_Origin","Country_of_Export",
    "Tariff_Number","Quantity","Quantity_UOM","Unit_Price","Total_Line_Value",
    "Net_Weight_KG","Gross_Weight_KG","Manufacturer_Name","Manufacturer_Address_1",
    "Manufacturer_Address_2","Manufacturer_City","Manufacturer_State","Manufacturer_Zip",
    "Manufacturer_Country","MID_Code","Buyer_Name","Buyer_Address_1","Buyer_Address_2",
    "Buyer_City","Buyer_State","Buyer_Zip","Buyer_Country","Buyer_ID_Number",
    "Consignee_Name","Consignee_Address_1","Consignee_Address_2","Consignee_City",
    "Consignee_State","Consignee_Zip","Consignee_Country","Consignee_ID_Number",
    "SICountry","SP1","SP2","Zone_Status","Privileged_Filing_Date","Line_Piece_Count",
    "ADD_Case_Number","CVD_Case_Number","AD_Non_Reimbursement_Statement",
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
    # Read starting at row 10 (skip first 9 rows)
    raw = pd.read_excel(path, skiprows=9, engine="openpyxl")

    # Detect presence of a commodity-description column (e.g. header contains "commodity")
    headers = [str(h).lower() for h in raw.columns]
    has_desc = any("commodity" in h for h in headers)

    # Build mapped dict, shifting columns if needed
    mapped = {}
    if has_desc:
        # Column C → index 2 → Commercial_Description
        mapped[HEADERS[xl_idx("C")]] = raw.iloc[:, 2]

        # For the rest, shift any src-index >= 2 right by one
        for src, tgt in MAPPING.items():
            src_idx = xl_idx(src)
            real_idx = src_idx + 1 if src_idx >= 2 else src_idx
            mapped[HEADERS[xl_idx(tgt)]] = raw.iloc[:, real_idx]
    else:
        # Old layout
        for src, tgt in MAPPING.items():
            mapped[HEADERS[xl_idx(tgt)]] = raw.iloc[:, xl_idx(src)]

    # Drop rows that are entirely blank
    tmp = pd.DataFrame(mapped).dropna(how="all").reset_index(drop=True)

    # Build the full frame + constants
    df = pd.DataFrame(index=tmp.index, columns=HEADERS)
    df.update(tmp)
    for tgt, val in CONSTANTS.items():
        df[HEADERS[xl_idx(tgt)]] = val

    # SG / HK overrides
    mid_c, cntry_c = HEADERS[xl_idx("T")], HEADERS[xl_idx("S")]
    sg_mask = (
        df[mid_c]
        .astype(str)
        .str.strip()
        .str.upper()
        .str.startswith("SG", na=False)
    )
    hk_mask = (
        df[mid_c]
        .astype(str)
        .str.strip()
        .str.upper()
        .str.startswith("HK", na=False)
    )
    df.loc[sg_mask, cntry_c] = "SG"
    df.loc[hk_mask, cntry_c] = "HK"

    # Preserve ZIP as 6-digit strings
    def _pad_zip(x):
        if pd.isna(x):
            return ""
        s = str(x).strip()
        m = re.fullmatch(r'(\d+)(?:\.0*)?', s)
        return f"{int(m.group(1)):06d}" if m else s

    for zip_col in ("Manufacturer_Zip", "Buyer_Zip"):
        if zip_col in df.columns:
            df[zip_col] = df[zip_col].apply(_pad_zip)

    return df


def save_chunks(
    df: pd.DataFrame,
    out_dir: str,
    mawb: str,
    rows: int = ROWS_PER_FILE
) -> int:
    os.makedirs(out_dir, exist_ok=True)

    # Choose suffix list based on rows-per-file
    if rows < 600:
        part_list = [f"{ltr}{i+1}" for ltr in string.ascii_uppercase for i in range(2)]
    else:
        part_list = list(string.ascii_uppercase)

    if len(df) > rows * len(part_list):
        raise ValueError("Too many rows for available file parts.")

    template_headers = [
        cell.value
        for cell in load_workbook(HEADER_PATH, read_only=True).active[1]
        if cell.value
    ]

    part = 0
    for start in range(0, len(df), rows):
        chunk = df.iloc[start : start + rows].copy()
        suffix = part_list[part]
        invoice = f"{mawb}-{suffix}"
        chunk["Invoice_No"] = invoice

        chunk = chunk.reindex(columns=template_headers)
        xlsx_path = os.path.join(out_dir, f"{invoice}.xlsx")
        with pd.ExcelWriter(xlsx_path, engine="xlsxwriter") as writer:
            chunk.to_excel(writer, index=False, header=False, startrow=1)
            workbook = writer.book
            worksheet = writer.sheets["Sheet1"]
            header_fmt = workbook.add_format(
                {
                    "bold": True,
                    "font_name": "Times New Roman",
                    "font_size": 12,
                    "align": "center",
                    "valign": "center",
                }
            )
            for idx, title in enumerate(template_headers):
                worksheet.write(0, idx, title, header_fmt)

            worksheet.set_column(0, 0, max(len(invoice), len("Invoice_No")) + 2)
            worksheet.set_column(5, 5, 20)

        part += 1

    return part
