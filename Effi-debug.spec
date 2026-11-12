# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
import sys


project_dir = Path(SPECPATH)
sys.path.insert(0, str(project_dir))
from effi_spec_common import collect_build_inputs


datas, binaries, hiddenimports, excludes = collect_build_inputs(project_dir)

a = Analysis(
    [str(project_dir / "application_extractor.py")],
    pathex=[str(project_dir)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
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
    name="Effi-debug",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
