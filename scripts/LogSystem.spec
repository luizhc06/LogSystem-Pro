# -*- mode: python ; coding: utf-8 -*-
# Build: scripts\build_exe.bat

from pathlib import Path

block_cipher = None
root = Path(SPECPATH).parent

datas = [
    (str(root / 'app.html'), '.'),
    (str(root / 'config.json'), '.'),
]

hiddenimports = [
    'webview',
    'openpyxl',
    'docx',
    'modules.database',
    'modules.bootstrap',
]

a = Analysis(
    [str(root / 'main.py')],
    pathex=[str(root)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
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
    name='LogSystemPro',
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
    icon=None,
)
