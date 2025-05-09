# -*- mode: python ; coding: utf-8 -*-
"""
Final spec – produces
dist\GA_broker_helper_<ver>\
├─ GA Broker Helper.lnk
└─ _internal\                    ← level-1
    ├─ GA Broker Helper.exe
    ├─ Resources\Logo\…
    └─ _internal\                ← level-2 (PyInstaller runtime)
        python310.dll …
"""
import os, shutil, sys
from PyInstaller.utils.hooks import collect_submodules
from PyInstaller.building.build_main import Analysis, PYZ, EXE, COLLECT

APP_VERSION = "1.4.3"
DIST_NAME   = f"GA_broker_helper_{APP_VERSION}"

# ── normal PyInstaller part ───────────────────────────────────────
block_cipher = None
a = Analysis(
    ["main_ui.py"],
    pathex=[],
    binaries=[],
    datas=[
        ("Resources/Logo/company_banner.png", "Resources/Logo"),
        ("Resources/Logo/company_logo.ico",   "Resources/Logo"),
        ("Resources/ExcelSplitter/Header Sample.xlsx",
                                          "Resources/ExcelSplitter"),
    ],
    hiddenimports=[
        "tkinterdnd2","customtkinter","pandas","openpyxl",
        "xlsxwriter","PyMuPDF",
    ] + collect_submodules("mini_updater"),
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)
exe = EXE(
    pyz, a.scripts, exclude_binaries=True,
    name="GA Broker Helper", icon="Resources/Logo/company_logo.ico",
    console=False,
)
coll = COLLECT(
    exe, a.binaries, a.zipfiles, a.datas,
    strip=False, upx=True, name=DIST_NAME,
)

# ── 3) post-build: wrap the original runtime ──────────────────────
dist_path   = os.path.join(os.getcwd(), "dist", DIST_NAME)
runtime_old = os.path.join(dist_path, "_internal")          # created by PyInstaller
outer_int   = os.path.join(dist_path, "_internal_wrap")     # temp name
exe_name    = "GA Broker Helper.exe"

# 3-A  make the wrapper folder
os.makedirs(outer_int, exist_ok=True)

# 3-B  move EXE and Resources into the wrapper
shutil.move(os.path.join(dist_path, exe_name), outer_int)
if os.path.isdir(os.path.join(dist_path, "Resources")):
    shutil.move(os.path.join(dist_path, "Resources"), outer_int)

# 3-C  move the original runtime folder *inside* the wrapper
shutil.move(runtime_old, os.path.join(outer_int, "_internal"))

# 3-D  rename wrapper to final '_internal'
final_int = os.path.join(dist_path, "_internal")
if os.path.exists(final_int):
    shutil.rmtree(final_int)
os.rename(outer_int, final_int)

# 3-E  create shortcut at dist root
try:
    import pythoncom
    from win32com.client import Dispatch
    pythoncom.CoInitialize()
    link   = os.path.join(dist_path, "GA Broker Helper.lnk")
    target = os.path.join(final_int, exe_name)
    sc     = Dispatch("WScript.Shell").CreateShortCut(link)
    sc.Targetpath, sc.WorkingDirectory, sc.IconLocation = target, final_int, target
    sc.save()
    pythoncom.CoUninitialize()
except Exception as e:
    print(f"[spec] shortcut failed: {e}", file=sys.stderr)
