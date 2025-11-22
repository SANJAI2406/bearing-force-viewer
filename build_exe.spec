# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec file for Bearing Force Viewer

import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# Collect easyocr data files and submodules
easyocr_datas = collect_data_files('easyocr')
easyocr_hiddenimports = collect_submodules('easyocr')

# Collect torch data
torch_datas = collect_data_files('torch')

a = Analysis(
    ['bearing_force_viewer.py'],
    pathex=[],
    binaries=[],
    datas=easyocr_datas + torch_datas,
    hiddenimports=[
        'easyocr',
        'torch',
        'torchvision',
        'PIL',
        'PIL.Image',
        'numpy',
        'matplotlib',
        'matplotlib.backends.backend_tkagg',
        'pandas',
        'openpyxl',
        'customtkinter',
        'cv2',
    ] + easyocr_hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='BearingForceViewer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # Keep console for debug output
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
