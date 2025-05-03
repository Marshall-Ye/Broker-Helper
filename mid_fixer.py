"""
mid_fixer.py  –  fixes numeric part of MID_Code (column T)
Adds helper columns:
    • MID_original  – original MID text
    • MID_flag      – True if numeric changed
No longer writes per‑chunk TXT; master report is handled by split_excel_core.
"""

from __future__ import annotations
import os, re, pandas as pd

# ---------- regex ladder ----------
_rx = lambda p: re.compile(p, re.I)
rules: list[tuple[re.Pattern, str | callable]] = [
    (_rx(r'\broom\s*([A-Za-z0-9\-]+)'),                    1),
    (_rx(r'\bunit\s*([A-Za-z0-9\-]+)'),                    2),
    (_rx(r'\b(?:shop|booth|stall)\s*([A-Za-z0-9\-]+)'),    3),
    (_rx(r'#\s*(\d+)[\-–](\d+)'),                          lambda m: m.group(1)+m.group(2)),
    (_rx(r'#\s*([A-Za-z0-9\-]+)'),                         5),
    (_rx(r'\bbuilding\s*(\d+)'),                           6),
    (_rx(r'\b(?:no\.?|road|rd\.?|street|st\.?|ave\.?|avenue)\s*(\d+)'), 7),
    (_rx(r'(\d+)'),                                        8),
]

def _clean(t:str)->str:
    t = re.sub(r'[A-Za-z\-]', '', t).lstrip('0') or '0'
    return t[:4]

def _valid(t:str)->bool: return t.isdigit() and 1 <= len(t) <= 4
def extract(addr:str|float):
    if not isinstance(addr,str): return None
    low = addr.lower()
    for rx,act in rules:
        m = rx.search(low)
        if not m: continue
        raw = m.group(1) if isinstance(act,int) else act(m)
        tok = _clean(raw)
        if _valid(tok): return tok
    return None

def mid_num(mid:str|float):
    if not isinstance(mid,str): return None
    m = re.search(r'(\d+)', mid)
    return m.group(1) if m else None

def repl_mid(mid:str,new:str)->str:
    return re.sub(r'\d+', new, mid, count=1)

EXEMPT = {"CNGUAJIA163GUA"}

# ---------- main fixer ----------
def fix_mids(df:pd.DataFrame, mawb:str, suffix:str, out_dir:str)->pd.DataFrame:
    df = df.copy()
    addr_c = df.columns[13]   # Manufacturer_Address_1
    mid_c  = df.columns[19]   # MID_Code

    df["MID_original"] = df[mid_c]

    addr_tok = df[addr_c].apply(extract)
    mid_tok  = df[mid_c].apply(mid_num)

    mask = (
        addr_tok.notna() & mid_tok.notna() &
        (addr_tok != mid_tok) &
        ~df[mid_c].isin(EXEMPT)
    )

    df.loc[mask, mid_c] = [
        repl_mid(old, new)
        for old, new in zip(df.loc[mask, mid_c], addr_tok[mask])
    ]

    df["MID_flag"] = mask
    return df
