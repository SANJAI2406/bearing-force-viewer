# -*- mode: python ; coding: utf-8 -*-
# Bearing Force Viewer - PyInstaller Spec File
# This spec ensures proper bundling of all dependencies

import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# Collect all matplotlib data files (fonts, styles, etc.)
matplotlib_datas = collect_data_files('matplotlib')

# Collect numpy data files
numpy_datas = collect_data_files('numpy')

# Hidden imports that PyInstaller might miss
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
    'numpy.random.common',
    'numpy.random.bounded_integers',
    'numpy.random.entropy',

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

    # PIL (optional but include if available)
    'PIL',
    'PIL.Image',
    'PIL.ImageTk',

    # CustomTkinter (optional)
    'customtkinter',

    # Packaging (needed by some dependencies)
    'packaging',
    'packaging.version',
    'packaging.specifiers',
    'packaging.requirements',
]

# Try to add customtkinter data files if available
try:
    ctk_datas = collect_data_files('customtkinter')
except:
    ctk_datas = []

# Combine all data files
all_datas = matplotlib_datas + numpy_datas + ctk_datas

a = Analysis(
    ['Bearing_force.py'],
    pathex=[],
    binaries=[],
    datas=all_datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Exclude heavy optional packages to reduce size
        'easyocr',
        'torch',
        'torchvision',
        'tensorflow',
        'scipy',
        'pandas',
        'IPython',
        'jupyter',
        'notebook',
    ],
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
    upx=False,  # Disable UPX compression - can cause PKG archive issues
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
