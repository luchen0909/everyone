# -*- mode: python ; coding: utf-8 -*-

import os


ROOT = os.path.abspath(SPECPATH)
BALANCE_SRC = os.path.abspath(os.path.join(ROOT, '..', '静态平衡表', 'src'))


a = Analysis(
    [os.path.join(ROOT, 'BOM查询工具.py')],
    pathex=[ROOT, BALANCE_SRC],
    binaries=[],
    datas=[],
    hiddenimports=['mrp_balance_tool', 'mrp_balance_tool.pipeline'],
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
    name='物料供需协同工具V2.0',
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

