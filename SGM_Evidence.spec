# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path


project_root = Path(SPECPATH)

a = Analysis(
    [str(project_root / "SGM_v1.1-de_clean_fixed.py")],
    pathex=[str(project_root)],
    binaries=[],
    datas=[
        (str(project_root / "assets" / "sgm_jokey.ico"), "assets"),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={"matplotlib": {"backends": ["TkAgg"]}},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="SGM Evidence 1.2.0",
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
    version=str(project_root / "version_info.txt"),
    icon=str(project_root / "assets" / "sgm_jokey.ico"),
)
