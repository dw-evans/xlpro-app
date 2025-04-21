# -*- mode: python ; coding: utf-8 -*-

import sys
import os

from PyInstaller.utils.hooks import copy_metadata
import PyInstaller.config

PyInstaller.config.CONF['workpath'] = "./build"
PyInstaller.config.CONF['distpath'] = "./dist"

datas = []
from pathlib import Path
def check_assets():
    checks = []
    checks_len = 4
    assets_dir = Path("assets")

    for p in assets_dir.glob("*"):
        print(f"checking {p}, {p.exists()}")

        if p.name == "xlpro.xlam":
            checks.append(True)

        elif p.name == "xlpro-cli.exe":
            checks.append(True)

        elif p.name == "config.toml": 
            checks.append(True)

        elif (
            p.name.lower().startswith("xlpro") and 
            p.name.lower().endswith(".whl")
        ):
            checks.append(True)
        else:
            pass


    print("asset checks completed successfully")

check_assets()

datas += [
    ("assets/", "assets/"),
]

sys.path.insert(0, os.getcwd())

import version

WIP_RELEASE = True

from pathlib import Path
EXE_NAME = version.VersionInfo.EXE_NAME

if WIP_RELEASE: 
    EXE_NAME = f"{Path(EXE_NAME).stem}_WIP{Path(EXE_NAME).suffix}"

VERSIONFILE = version.create_version_file_get_path()

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[
        'xlpro_installer',
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
    version=VERSIONFILE,
)
