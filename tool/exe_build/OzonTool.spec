# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

BUILD_ROOT = Path(SPECPATH).resolve().parent
TOOL_ROOT = BUILD_ROOT.parent
APP_ENTRY = TOOL_ROOT / "app.py"

a = Analysis(
    [str(APP_ENTRY)],
    pathex=[str(TOOL_ROOT)],
    binaries=[],
    datas=[(str(TOOL_ROOT / "builtin_mockups"), "builtin_mockups")],
    hiddenimports=[
        "image_providers",
        "text_providers",
        "config_store",
        "image_api_client",
        "scene_generator",
        "local_scene_composer",
        "mockup_template",
        "batch_upload",
        "batch_upload.core",
        "batch_upload.video_ops",
        "batch_upload.inventory_ops",
        "openpyxl",
        "openpyxl.cell",
        "openpyxl.styles",
        "openpyxl.workbook",
    ],
    hookspath=[],
    hooksconfig={},
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
    name="OzonTool",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
