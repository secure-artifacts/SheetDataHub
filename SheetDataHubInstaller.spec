# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

root = Path(SPECPATH)

a = Analysis(
    [str(root / 'installer.py')],
    pathex=[str(root)],
    binaries=[],
    datas=[
        (str(root / 'assets' / 'app.ico'), 'assets'),
        (str(root / 'dist-onedir' / 'SheetDataHub'), 'payload/SheetDataHub'),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[str(root / 'packaging' / 'pyi_rth_skip_numpy.py')],
    excludes=['numpy', 'pandas', 'matplotlib', 'scipy', 'tkinter'],
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
    name='SheetDataHub-Setup',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    icon=[str(root / 'assets' / 'app.ico')],
)
