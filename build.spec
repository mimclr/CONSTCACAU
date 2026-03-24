# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'openpyxl',
        'openpyxl.workbook',
        'openpyxl.worksheet',
        'openpyxl.reader',
        'openpyxl.utils',
        'tkinter',
        'tkinter.filedialog',
        'tkinter.ttk',
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'numpy', 'scipy', 'PIL', 'cv2',
        'pandas', 'IPython', 'jupyter',
    ],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Analise_Consistencia_LaborRural',
    debug=False,
    strip=False,
    upx=False,
    console=False,
    runtime_tmpdir=None,
)
