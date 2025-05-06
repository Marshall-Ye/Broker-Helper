# ga_broker_helper.spec
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main_ui.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('Resources/Logo/company_banner.png', 'Resources/Logo'),
        ('Resources/Logo/company_logo.ico',   'Resources/Logo'),
        ('Resources/ExcelSplitter/Header Sample.xlsx', 'Resources/ExcelSplitter'),
    ],
    hiddenimports=['tkinterdnd2', 'customtkinter'],
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    exclude_binaries=True,
    name='GA Broker Helper V1.3',
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
    name='GA Broker Helper V1.3',
)
