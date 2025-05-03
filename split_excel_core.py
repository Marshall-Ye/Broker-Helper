"""
split_excel_core.py
-------------------
• Builds DataFrame from source workbook
• Splits into N‑row chunks
• Runs mid_fixer.fix_mids() on each chunk
• Writes Excel files and master MID_report_<MAWB>.txt
"""

import os, sys, pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
import mid_fixer

# ---------- mapping / constants ----------
MAPPING   = {"B":"F","D":"G","F":"J","E":"L","G":"M","H":"N","I":"P","K":"R","M":"T"}
CONSTANTS = {"D":"CN","E":"CN","S":"CN","H":"PCS"}

# ---------- header list ----------
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

# ---------- helpers ----------
def xl_idx(col:str)->int:
    idx = 0
    for ch in col.upper():
        idx = idx * 26 + ord(ch) - 64
    return idx - 1

def get_mawb(path:str)->str:
    wb = load_workbook(path, read_only=True, data_only=True)
    val = str(wb.active["U9"].value).strip()
    wb.close()
    return val

# ---------- build DataFrame (without MID fix) ----------
def prepare_dataframe(path:str)->pd.DataFrame:
    raw = pd.read_excel(path, skiprows=9, engine="openpyxl")
    mapped = {HEADERS[xl_idx(t)]: raw.iloc[:, xl_idx(s)] for s, t in MAPPING.items()}
    tmp = pd.DataFrame(mapped).dropna(how="all").reset_index(drop=True)

    df = pd.DataFrame(index=tmp.index, columns=HEADERS)
    df.update(tmp)

    for tgt, val in CONSTANTS.items():
        df[HEADERS[xl_idx(tgt)]] = val

    # SG override
    mid_col, cntry_col = HEADERS[xl_idx("T")], HEADERS[xl_idx("S")]
    sg_mask = df[mid_col].astype(str).str.strip().str.upper().str.startswith("SG")
    df.loc[sg_mask, cntry_col] = "SG"

    return df

# ---------- save chunks ----------
def save_chunks(df:pd.DataFrame, out_dir:str, mawb:str, rows:int=998) -> int:
    os.makedirs(out_dir, exist_ok=True)
    letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    font, align = Font(name="Calibri", size=11), Alignment(horizontal="left")

    master = os.path.join(out_dir, f"MID_report_{mawb}.txt")
    with open(master, "w", encoding="utf-8") as rpt:
        rpt.write("File\tRow\tOriginalMID\tFinalMID\tManufacturer\tAddress\n")

        part = 0
        for start in range(0, len(df), rows):
            chunk = df.iloc[start:start + rows].copy()
            suffix = letters[part]
            invoice = f"{mawb}-{suffix}"
            chunk["Invoice_No"] = invoice

            # --- fix MIDs ---
            chunk = mid_fixer.fix_mids(chunk, mawb, suffix, out_dir)

            # --- write workbook ---
            xlsx = os.path.join(out_dir, f"{invoice}.xlsx")
            chunk.to_excel(xlsx, index=False)

            wb = load_workbook(xlsx)
            ws = wb.active
            for c in ws[1]:
                c.font = font
                c.alignment = align
            ws.column_dimensions["A"].width = max(len(invoice), len("Invoice_No")) + 2
            ws.column_dimensions["F"].width = 20

            # --- append to master TXT report ---
            for r, (_, row) in enumerate(chunk.iterrows(), start=2):
                if row["MID_flag"]:
                    rpt.write(f"{invoice}.xlsx\t{r}\t{row['MID_original']}\t"
                              f"{row['MID_Code']}\t{row['Manufacturer_Name']}\t"
                              f"{row['Manufacturer_Address_1']}\n")

            # --- drop helper columns (safe order) ---
            helpers  = [h for h in ("MID_flag", "MID_original") if h in chunk.columns]
            col_idxs = [chunk.columns.get_loc(h) + 1 for h in helpers]   # Excel is 1‑based
            for col in sorted(col_idxs, reverse=True):                   # delete highest first
                if col <= ws.max_column:
                    ws.delete_cols(col)

            wb.save(xlsx)
            wb.close()
            part += 1

    return part

# ---------- CLI ----------
if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python split_excel_core.py <input.xlsx> <output_folder>")
        sys.exit(1)

    src, dst = sys.argv[1], sys.argv[2]
    mawb_root = get_mawb(src)
    parts = save_chunks(prepare_dataframe(src), dst, mawb_root)
    print(f"Done – {parts} file(s) written to '{dst}'.")
