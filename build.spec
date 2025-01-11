# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('gui', 'gui'),  # Copy thư mục gui
        ('samples', 'samples'),  # Copy thư mục samples nếu cần
    ],
    hiddenimports=[
        'pandas',
        'pandas._libs.tslibs.base',
        'pandas._libs.tslibs.offsets',
        'pandas._libs.tslibs.parsing',
        'pandas._libs.tslibs.period',
        'pandas._libs.tslibs.strptime',
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.timestamps',
        'pandas._libs.tslibs.timezones',
        'pandas._libs.tslibs.vectorized',
        'numpy',
        'docx',
        'bs4',
        'jinja2',
        'tkinter',
        'email.mime.text',
        'email.mime.multipart',
        'email.mime.image',
        'email.mime.application',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'scipy',
        'PIL',
        'torch',
        'tensorflow'
    ],
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
    name='MergeMail',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,  # Tắt UPX compression
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # False để ẩn console window
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None
) 