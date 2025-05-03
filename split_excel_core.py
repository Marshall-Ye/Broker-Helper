# split_excel_core.py
"""
split_excel_core.py
-------------------
• Reads source workbook
• Maps / copies columns
• Splits data into N-row Excel files
• Uses customs-approved header sample
• Preserves leading zeros in ZIP codes (exactly 6 digits, no decimals)
"""

import os
import sys
import re
import string
import pandas as pd
from openpyxl import load_workbook

ROWS_PER_FILE = 998

MAPPING = {
    "B": "F", "D": "G", "F": "J", "E": "L",
    "G": "M", "H": "N", "I": "P", "K": "R", "M": "T"
}

CONSTANTS = {
    "D": "CN",   # Country_of_Origin
    "E": "CN",   # Country_of_Export
    "S": "CN",   # Manufacturer_Country
    "H": "PCS"   # Quantity_UOM
}

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

    # 1) Map & drop blank rows
    mapped = {
        HEADERS[xl_idx(tgt)]: raw.iloc[:, xl_idx(src)]
        for src, tgt in MAPPING.items()
    }
    tmp = pd.DataFrame(mapped).dropna(how="all").reset_index(drop=True)

    # 2) Build full frame + constants
    df = pd.DataFrame(index=tmp.index, columns=HEADERS)
    df.update(tmp)
    for tgt, val in CONSTANTS.items():
        df[HEADERS[xl_idx(tgt)]] = val

    # 3) SG override
    mid_c, cntry_c = HEADERS[xl_idx("T")], HEADERS[xl_idx("S")]
    sg_mask = (
        df[mid_c]
        .astype(str)
        .str.strip()
        .str.upper()
        .str.startswith("SG", na=False)
    )
    df.loc[sg_mask, cntry_c] = "SG"

    # 4) Preserve ZIP codes as 6-digit strings
    def _pad_zip(x):
        if pd.isna(x):
            return ""
        s = str(x).strip()
        # Recognize pure number or floats ending in .0
        m = re.fullmatch(r'(\d+)(?:\.0*)?', s)
        if m:
            num = int(m.group(1))
            return f"{num:06d}"
        # Already text or contains letters — leave as is
        return s

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

    # Generate A1, A2, B1, B2, … Z1, Z2
    part_list = [f"{ltr}{i+1}" for ltr in string.ascii_uppercase for i in range(2)]

    # Load customs-approved header row
    sample_path = os.path.join(os.path.dirname(__file__),
                               "RealData", "Header Sample.xlsx")
    sample_wb   = load_workbook(sample_path, read_only=True)
    sample_ws   = sample_wb.active
    template_headers = [cell.value for cell in sample_ws[1] if cell.value]
    sample_wb.close()

    part = 0
    for start in range(0, len(df), rows):
        chunk  = df.iloc[start:start + rows].copy()
        suffix = part_list[part]
        invoice = f"{mawb}-{suffix}"
        chunk["Invoice_No"] = invoice

        # Reorder to match the template exactly
        chunk = chunk.reindex(columns=template_headers)

        # Write via XlsxWriter
        xlsx_path = os.path.join(out_dir, f"{invoice}.xlsx")
        with pd.ExcelWriter(xlsx_path, engine="xlsxwriter") as writer:
            chunk.to_excel(writer, index=False, header=False, startrow=1)

            workbook  = writer.book
            worksheet = writer.sheets["Sheet1"]

            header_fmt = workbook.add_format({
                'bold': True,
                'font_name': 'Times New Roman',
                'font_size': 12,
                'align': 'center',
                'valign': 'center'
            })

            # Write header row at row 0
            for col_idx, title in enumerate(template_headers):
                worksheet.write(0, col_idx, title, header_fmt)

            # Column widths
            worksheet.set_column(0, 0, max(len(invoice), len("Invoice_No")) + 2)
            worksheet.set_column(5, 5, 20)

        part += 1

    return part

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python split_excel_core.py <input.xlsx> <output_folder>")
        sys.exit(1)
    src, dst = sys.argv[1], sys.argv[2]
    parts = save_chunks(prepare_dataframe(src), dst, get_mawb(src))
    print(f"Done – {parts} file(s) written to '{dst}'.")
