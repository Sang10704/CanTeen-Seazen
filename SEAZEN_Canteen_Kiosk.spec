# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller Spec file for SEAZEN Canteen Kiosk Standalone Executable
Đóng gói 1 file EXE duy nhất chứa đầy đủ HTML, âm thanh MP3, Logo và Biểu tượng.
Không nhúng token cứng; token được quản lý động và mã hóa DPAPI tại %LOCALAPPDATA%.
"""

import os
import sys

block_cipher = None

a = Analysis(
    ['Canteen_Kiosk_Runner.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('Canteen_Kiosk_Production.html', '.'),
        ('mp3', 'mp3'),
        ('seazen_logo.png', '.'),
        ('canteen_icon.ico', '.')
    ],
    hiddenimports=[
        'openpyxl',
        'openpyxl.styles',
        'msal',
        'flask',
        'requests',
        'ctypes',
        'ctypes.wintypes'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
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
