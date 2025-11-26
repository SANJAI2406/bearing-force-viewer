# -*- mode: python ; coding: utf-8 -*-
# Bearing Force Viewer - PyInstaller Spec File (FULL VERSION with OCR)
# This spec includes ALL modules including easyocr and torch

import sys
import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules, collect_all

block_cipher = None

# Collect all data files from dependencies
matplotlib_datas = collect_data_files('matplotlib')
numpy_datas = collect_data_files('numpy')

# Try to collect easyocr data files (OCR models)
try:
    easyocr_datas = collect_data_files('easyocr')
except:
    easyocr_datas = []

# Try to collect torch data
try:
    torch_datas = collect_data_files('torch')
except:
    torch_datas = []

# Try to add customtkinter data files
try:
    ctk_datas = collect_data_files('customtkinter')
except:
    ctk_datas = []

# Hidden imports - include EVERYTHING
hidden_imports = [
    # Matplotlib backends
    'matplotlib',
    'matplotlib.pyplot',
    'matplotlib.figure',
    'matplotlib.backends.backend_tkagg',
    'matplotlib.backends.backend_agg',
    'matplotlib.backend_bases',

    # Numpy
    'numpy',
    'numpy.core._methods',
    'numpy.lib.format',
    'numpy.random',

    # Tkinter
    'tkinter',
    'tkinter.ttk',
    'tkinter.filedialog',
    'tkinter.messagebox',

    # Standard library
    'json',
    'pathlib',
    'concurrent.futures',
    'subprocess',
    'platform',
    're',

    # PIL
    'PIL',
    'PIL.Image',
    'PIL.ImageTk',

    # CustomTkinter
    'customtkinter',

    # Packaging
    'packaging',
    'packaging.version',
    'packaging.specifiers',
    'packaging.requirements',

    # EasyOCR and dependencies
    'easyocr',
    'easyocr.easyocr',
    'easyocr.detection',
    'easyocr.recognition',
    'easyocr.utils',

    # PyTorch (required by easyocr)
    'torch',
    'torch.nn',
    'torch.nn.functional',
    'torchvision',
    'torchvision.transforms',

    # Other easyocr dependencies
    'cv2',
    'scipy',
    'scipy.ndimage',
    'skimage',
    'skimage.transform',
]

# Collect all submodules for critical packages
hidden_imports += collect_submodules('easyocr')
hidden_imports += collect_submodules('torch')
hidden_imports += collect_submodules('torchvision')

# Combine all data files
all_datas = matplotlib_datas + numpy_datas + ctk_datas + easyocr_datas + torch_datas

a = Analysis(
    ['Bearing_force.py'],
    pathex=[],
    binaries=[],
    datas=all_datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],  # NO EXCLUSIONS - include everything
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Bearing_Force',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,  # Disable UPX compression - prevents PKG archive corruption
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUI application, no console
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
