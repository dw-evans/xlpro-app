

import datetime

class VersionInfo:
    VERSION = "0.1.0.0"
    NAME = "xlpro-installer"
    EXE_NAME = f"{NAME}-{VERSION}.exe"
    AUTHOR = "D.Evans"
    DATE = datetime.datetime.now().strftime("%Y-%m-%d")
    DESCRIPTION_SHORT = "xlpro Windows installer (BETA)"
    COMPANY_NAME="Dash Software"

import pyinstaller_versionfile

from pathlib import Path
wd = Path(__file__).parent

def create_version_file_get_path() -> str:
    pyinstaller_versionfile.create_versionfile(
        output_file=(ret:=(wd / "version.txt")),
        version=VersionInfo.VERSION,
        company_name="Dash Software",
        file_description=VersionInfo.DESCRIPTION_SHORT,
        internal_name=VersionInfo.NAME,
        legal_copyright=f"© {VersionInfo.AUTHOR}. All rights reserved.",
        original_filename=VersionInfo.EXE_NAME,
        product_name=VersionInfo.NAME,
        # translations=[0, 1200]
    )
    return str(ret)

pass