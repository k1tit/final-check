# -*- mode: python ; coding: utf-8 -*-
# PyInstaller: папка dist\PF_BP_PY_ZY_Web\ — запуск PF_BP_PY_ZY_Web.exe
# Сборка: build_portable.cmd

block_cipher = None

a = Analysis(
    ["run_checks_web.py"],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=["build_checks"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="PF_BP_PY_ZY_Web",
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
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="PF_BP_PY_ZY_Web",
)
