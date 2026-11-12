# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
import sys

from PyInstaller.utils.hooks import collect_data_files


project_dir = Path(SPECPATH)
sys.path.insert(0, str(project_dir))
from effi_spec_common import collect_build_inputs


# 沿用单文件版的依赖、模型、模板和发行包元数据收集。
datas, binaries, hiddenimports, excludes = collect_build_inputs(project_dir)
datas += collect_data_files("customtkinter", includes=["assets/**/*"])

# 离线交付必须带齐两套模型，不在运行时补下载。
for model_name in ("PP-OCRv5_mobile_det", "PP-OCRv5_mobile_rec"):
    for filename in ("inference.json", "inference.pdiparams", "inference.yml"):
        model_file = project_dir / ".paddlex" / "official_models" / model_name / filename
        if not model_file.is_file():
            raise SystemExit("本地 OCR 模型文件不完整，无法构建离线发布版本")

a = Analysis(
    [str(project_dir / "application_extractor.py")],
    pathex=[str(project_dir)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[str(project_dir / "effi_release_runtime_hook.py")],
    excludes=excludes,
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

# 依赖交给 COLLECT，EXE 启动时不再整体解包。
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="Effi",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    contents_directory="_internal",
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="Effi",
)
