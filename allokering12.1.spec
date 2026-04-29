# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller build spec for Allokering."""
from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files


project_root = Path(SPECPATH)
tkdnd_datas = collect_data_files("tkinterdnd2")
extra_datas = [
    (str(project_root / "wms_sok79.py"), "."),
    (str(project_root / "packaging" / "windows" / "app.ico"), "."),
]

a = Analysis(
    ["allokering12.1.py"],
    pathex=[str(project_root)],
    binaries=[],
    datas=tkdnd_datas + extra_datas,
    hiddenimports=["tkinterdnd2"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "pytest",
        "tests",
        "webapp",
        "ask",
        "syntetiskdata",
        "testdata",
        "matplotlib",
        "PyQt5",
        "PyQt6",
        "PySide2",
        "PySide6",
        "pygame",
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="Allokering",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(project_root / "packaging" / "windows" / "app.ico"),
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="Allokering",
)
