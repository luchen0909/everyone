# -*- mode: python ; coding: utf-8 -*-

import importlib.util
import os


ROOT = os.path.abspath(SPECPATH)
BALANCE_SRC = os.path.abspath(os.path.join(ROOT, '..', '静态平衡表', 'src'))
OPTIONAL_IMPORTS = ['pyodbc'] if importlib.util.find_spec('pyodbc') else []


a = Analysis(
    [os.path.join(ROOT, 'BOM查询工具.py')],
    pathex=[ROOT, BALANCE_SRC],
    binaries=[],
    datas=[
        (os.path.join(ROOT, 'assets', '静态平衡表模板.xlsx'), 'assets'),
    ],
    hiddenimports=['mrp_balance_tool', 'mrp_balance_tool.pipeline', *OPTIONAL_IMPORTS],
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
    [],
    exclude_binaries=True,
    name='SupplyCoordinationToolV3',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name='SupplyCoordinationToolV3',
)

app = BUNDLE(
    coll,
    name='SupplyCoordinationToolV3.app',
    icon=None,
    bundle_identifier='com.luchen.supplycoordination.v3',
)
