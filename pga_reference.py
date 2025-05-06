"""
pga_reference.py  –  PGA Reference Tab (4-column grid)
------------------------------------------------------
Shows hard-coded PGA blocks. First line (AL1, AQ1, …) is big & bold.
Blocks are arranged in up-to-4 columns to use horizontal space.
"""

import os, sys
import customtkinter as ctk

# ---------- reference data (edit here) ----------
REF_BLOCKS = [
    ["AL1", "AL2", "FD2", "FD3", "Delete"],
    ["AQ1", "APHIS", "APQ", "A"],
    ["FW1", "FWS", "FWS", "E"],
    ["FD1", "FDA", "FOO", "A"],
    ["EP7", "EPA", "TS1", "A"],
]

TITLE_FONT  = ("Arial", 14, "bold")
LINE_FONT   = ("Arial", 12)
CODE_FONT   = ("Arial", 16, "bold")
TEXT_COLOR  = "#cccccc"
COLUMNS     = 4                     # <— number of grid columns


class PGAReferenceTab:
    def __init__(self, parent):
        # Scrollable area
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=10, pady=10)

        # configure grid columns inside scroll frame
        for c in range(COLUMNS):
            scroll.grid_columnconfigure(c, weight=1, uniform="col")

        for idx, block in enumerate(REF_BLOCKS):
            if not block:
                continue

            row = idx // COLUMNS
            col = idx % COLUMNS

            cell = ctk.CTkFrame(scroll, fg_color="transparent")
            cell.grid(row=row, column=col, sticky="nw", padx=10, pady=8)

            # header code
            ctk.CTkLabel(cell, text=block[0], font=CODE_FONT).pack(anchor="w")

            # body lines
            for line in block[1:]:
                ctk.CTkLabel(cell, text=line, font=LINE_FONT,
                             text_color=TEXT_COLOR).pack(anchor="w")
