# -*- mode: python ; coding: utf-8 -*-
# PyInstaller build recipe: one Windows EXE, no console window.
#   pyinstaller --noconfirm --clean ttron2excel.spec
# Optional: put an "icon.ico" next to this file to give the EXE an icon.
import os

APP_NAME = "TTRON2Excel"

a = Analysis(
    ["ttron2excel_app.py"],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        "openpyxl",                            # pandas loads it only when writing .xlsx
        "matplotlib.backends.backend_agg",
        "matplotlib.backends.backend_pdf",
        "matplotlib.backends.backend_tkagg",   # review window
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[
        "PyQt5", "PyQt6", "PySide2", "PySide6",   # not used; keeps the EXE smaller
        "IPython", "jupyter_client", "notebook", "pytest",
    ],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,           # UPX-packed EXEs are flagged by antivirus software more often
    console=False,       # GUI application: no black console window
    runtime_tmpdir=None,
    icon="icon.ico" if os.path.exists("icon.ico") else None,
)
