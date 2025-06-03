# ga_broker_helper.spec  –  default onedir layout
import os
from PyInstaller.utils.hooks import collect_submodules
from PyInstaller.building.build_main import Analysis, PYZ, EXE, COLLECT

APP_VERSION = "1.5.3"
DIST_NAME   = f"GA_broker_helper_{APP_VERSION}"

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
COLLECT(
    exe, a.binaries, a.zipfiles, a.datas,
    strip=False, upx=True, name=DIST_NAME,
)
