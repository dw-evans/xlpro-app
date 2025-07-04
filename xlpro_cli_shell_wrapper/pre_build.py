import datetime
import xlpro


class VersionInfo:
    VERSION = xlpro.__version__
    NAME = "xlpro"
    EXE_NAME = f"{NAME}-v{VERSION}-installer.exe"
    AUTHOR = "D Evans"
    DATE = datetime.datetime.now().strftime("%Y-%m-%d")
    DESCRIPTION_SHORT = "xlpro Windows installer"
    COMPANY_NAME = "D Evans"


import pyinstaller_versionfile

from pathlib import Path

wd = Path(__file__).parent


def create_version_file_get_path() -> str:
    pyinstaller_versionfile.create_versionfile(
        output_file=(ret := (wd / "version.txt")),
        version=VersionInfo.VERSION + ".0",
        company_name=VersionInfo.COMPANY_NAME,
        file_description=VersionInfo.DESCRIPTION_SHORT,
        internal_name=VersionInfo.NAME,
        legal_copyright=f"© {VersionInfo.AUTHOR}. All rights reserved.",
        original_filename=VersionInfo.EXE_NAME,
        product_name=VersionInfo.NAME,
    )
    return str(ret)


pass

import json
from pathlib import Path

wd = Path(__file__).parent


def create_winres_json():
    d = {
        "RT_GROUP_ICON": {
            "APP": {
                "0000": [
                    # "favicon.ico",
                    "icon.png",
                    "icon16.png",
                ]
            }
        },
        "RT_MANIFEST": {
            "#1": {
                "0409": {
                    "identity": {"name": "", "version": ""},
                    "description": "",
                    "minimum-os": "win10",
                    "execution-level": "as invoker",
                    "ui-access": False,
                    "auto-elevate": False,
                    "dpi-awareness": "system",
                    "disable-theming": False,
                    "disable-window-filtering": False,
                    "high-resolution-scrolling-aware": False,
                    "ultra-high-resolution-scrolling-aware": False,
                    "long-path-aware": False,
                    "printer-driver-isolation": False,
                    "gdi-scaling": False,
                    "segment-heap": False,
                    "use-common-controls-v6": False,
                }
            }
        },
        "RT_VERSION": {
            "#1": {
                "0000": {
                    "fixed": {
                        "file_version": f"{xlpro.__version__}.0",
                        "product_version": f"{xlpro.__version__}",
                    },
                    "info": {
                        "0409": {
                            "Comments": "",
                            "CompanyName": "",
                            "FileDescription": "xlpro-server",
                            "FileVersion": f"{xlpro.__version__}",
                            "InternalName": "xlpro-server",
                            "LegalCopyright": "Copyright 2025 (c) Daniel Evans",
                            "LegalTrademarks": "",
                            "OriginalFilename": "",
                            "PrivateBuild": "",
                            "ProductName": "",
                            "ProductVersion": "",
                            "SpecialBuild": "",
                        }
                    },
                }
            }
        },
    }

    # d = {
    #     "Version": {
    #         "CompanyName": "xlpro",
    #         "FileDescription": "xlpro Shell Wrapper",
    #         "FileVersion": f"{xlpro.__version__}.0",
    #         "ProductVersion": f"{xlpro.__version__}",
    #         "ProductName": "xlpro-server",
    #     },
    #     "IconPath": "favicon.ico",
    # }
    fout = wd / "winres/winres.json"
    wd.parent.mkdir(exist_ok=True)
    with open(fout, "w", encoding="utf-8") as f:
        f.write(json.dumps(d, indent=4))

    # with open(fout, "rb") as f:
    #     start = f.read(3)
    #     print(start)  # b'\xef\xbb\xbf' → means BOM exists


if __name__ == "__main__":
    create_winres_json()
