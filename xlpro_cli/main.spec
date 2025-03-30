# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import copy_metadata
import PyInstaller.config

PyInstaller.config.CONF['workpath'] = "./build"
PyInstaller.config.CONF['distpath'] = "./dist"

datas = []
# datas += [
#     ('utils.py', '.'),
#     ('commands.py', '.'),
# ]
datas += copy_metadata('readchar', recursive=True)

EXE_NAME = "xlpro-cli"

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[
        'xlpro_cli',
        'readchar',
        'rich',
        'regex',
        'uuid',

    ],
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
    name=EXE_NAME,
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
    onefile=True,
)
