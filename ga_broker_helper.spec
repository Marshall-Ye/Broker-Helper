# -*- mode: python ; coding: utf-8 -*-
#  Re-spun to the original state – no updater-related tweaks,
#  no additional cffi collection.  Folder (COLLECT) build as before.

block_cipher = None

a = Analysis(
    ['main_ui.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('Resources/Logo/company_banner.png',              'Resources/Logo'),
        ('Resources/Logo/company_logo.ico',                 'Resources/Logo'),
        ('Resources/ExcelSplitter/Header Sample.xlsx',      'Resources/ExcelSplitter'),
    ],
    hiddenimports=[
        'tkinterdnd2',
        'customtkinter',
        'pandas',
        'openpyxl',
        'xlsxwriter',
        'PyMuPDF',
    ],
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    exclude_binaries=True,
    name='GA Broker Helper',                # original exe name/folder
    icon='Resources/Logo/company_logo.ico',
    console=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    name='GA Broker Helper',                # output folder under dist/
)
