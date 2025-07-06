# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import copy_metadata
import PyInstaller.config

PyInstaller.config.CONF['workpath'] = "./build"
PyInstaller.config.CONF['distpath'] = "./dist"

datas = []
datas += copy_metadata('readchar', recursive=True)
datas += [("cli-header.txt", "xlpro_cli")]
datas += [("cli-header.txt", ".")]

import sys
sys.path.insert(0, os.getcwd())

import pre_build
pre_build.main()

import version

EXE_NAME = version.VersionInfo.EXE_NAME
VERSIONFILE = version.create_version_file_get_path()

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
    icon="favicon.ico",
    # icon=r"C:\Users\Daniel Evans\projects\xlpro\xlpro_installer\favicon.ico"
)
