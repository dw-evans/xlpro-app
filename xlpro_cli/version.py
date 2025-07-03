

import datetime
import xlpro

xlpro.__version__

__version__ = f"{xlpro.__version__}.0"

class VersionInfo:
    VERSION = __version__
    NAME = "xlpro-cli"
    EXE_NAME = f"{NAME}.exe"
    AUTHOR = "D. Evans"
    DATE = datetime.datetime.now().strftime("%Y-%m-%d")
    DESCRIPTION_SHORT = "xlpro Windows Installer"
    COMPANY_NAME="D. Evans"

import pyinstaller_versionfile

from pathlib import Path
wd = Path(__file__).parent

def create_version_file_get_path() -> str:
    pyinstaller_versionfile.create_versionfile(
        output_file=(ret:=(wd / "version.txt")),
        version=VersionInfo.VERSION,
        company_name=VersionInfo.COMPANY_NAME,
        file_description=VersionInfo.DESCRIPTION_SHORT,
        internal_name=VersionInfo.NAME,
        legal_copyright=f"Copyright (c) 2025 Daniel Evans. All rights reserved.",
        original_filename=VersionInfo.EXE_NAME,
        product_name=VersionInfo.NAME,
    )
    return str(ret)

pass