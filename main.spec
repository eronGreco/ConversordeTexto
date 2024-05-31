# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=['C:\\Users\\conta\\PycharmProjects\\extracttext'],
    binaries=[],
    datas=[
        ('C:\\Users\\conta\\PycharmProjects\\extracttext\\tkdnd2.9.4', 'tkdnd2.9.4'),
    ],
    hiddenimports=[
        'tkinterdnd2',
        'PyPDF2',
        'docx',  # python-docx
        'pptx',  # python-pptx
        'pytesseract',
        'PIL',  # Pillow
        'fitz',  # pymupdf
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name='main',
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

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='main',
)
