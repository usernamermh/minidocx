# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

project_root = Path.cwd()

a = Analysis(
    [str(project_root / 'tray_app.py')],
    pathex=[str(project_root)],
    binaries=[],
    datas=[
        (str(project_root / 'static'), 'static'),
        (str(project_root / 'tray_icon.png'), '.'),
        (str(project_root / 'tray_icon.ico'), '.'),
    ],
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
    name='MiniDocxTray',
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
    icon=[str(project_root / 'tray_icon.ico')],
)
