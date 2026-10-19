# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_dynamic_libs, collect_submodules

qt_hiddenimports = collect_submodules('PySide6') + collect_submodules('shiboken6')
qt_binaries = collect_dynamic_libs('PySide6') + collect_dynamic_libs('shiboken6')

a = Analysis(
    ['Experimental_Baukasten.py'],
    pathex=[],
    binaries=qt_binaries,
    datas=[('procedures.json', '.'), ('solvents.json', '.'), ('trivial_names.json', '.'), ('reaction_learning.json', '.')],
    hiddenimports=qt_hiddenimports + ['shiboken6.Shiboken'],
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
    name='Experimental_Baukasten',
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
