# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['Canteen_Kiosk_Runner.py'],
    pathex=[],
    binaries=[],
    datas=[('Canteen_Kiosk_Production.html', '.'), ('token_cache.bin', '.'), ('mp3', 'mp3')],
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
    name='SEAZEN_Canteen_Kiosk',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['canteen_icon.ico'],
)
