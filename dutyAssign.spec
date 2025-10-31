# -*- mode: python ; coding: utf-8 -*-

import os
from PyInstaller.utils.hooks import get_package_paths

# 'ortools'パッケージのパスを取得
_, ortools_pkg_dir = get_package_paths('ortools')
# 'ortools.dll'へのフルパスを構築
# ortools_dll_path = os.path.join(ortools_pkg_dir, '.libs', 'ortools.dll')

a = Analysis(
    ['dutyAssign.py'],
    pathex=[],
    datas=[],
    hiddenimports=['cp_model'],
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
    name='dutyAssign',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
