# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['Zamini_Converter_v1.0.0.py'],
    pathex=[],
    binaries=[],
    datas=[('tcl/tcl8.6', 'tcl/tcl8.6'), ('tcl/tk8.6', 'tcl/tk8.6'), ('tcl/tkdnd2.8', 'tcl/tkdnd2.8')],
    hiddenimports=[],
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
    name='Zamini_Converter_v1.0.0',
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
    icon=['Zamini_Musafir_logo.ico'],
)
