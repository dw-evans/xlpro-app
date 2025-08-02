# Builds all code and moves it to the assets dir


from pathlib import Path
import subprocess
import shutil

import xlpro

wd = Path(__file__).parent


paths = [
    (
        "startfiles", 
        "startfiles",
    ),
    (
        f"../xlpro_module/dist/xlpro-{xlpro.__version__}-py3-none-any.whl", 
        f"src/xlpro-{xlpro.__version__}-py3-none-any.whl",
    ),
    (
        f"../xlpro_module/config_production.toml",
        "config.toml"
    ),
    (
        f"../xlpro_addin/dist/xlpro.xlam",
        f"src/xlpro.xlam",
    ),
    (
        f"../xlpro_cli/dist/xlpro-cli.exe",
        "xlpro-cli.exe",
    ),
    (
        f"../xlpro_examples",
        f"examples",
    ),
    (
        f"../LICENSE",
        f"LICENSE.txt",
    ),
    (
        f"../licenses/uv_LICENSE",
        f"licenses/uv_LICENSE.txt",
    ),
    (
        f"../licenses/third_party_licenses_summary.md",
        f"licenses/third_party_licenses_summary.txt",
    ),
    (
        f"../licenses/third_party_licenses.md",
        f"licenses/third_party_licenses.txt",
    ),
    (
        f"../xlpro_cli_shell_wrapper/xlpro-server.exe",
        f"xlpro-server.exe",
    )
]

import os
import shutil

import xlpro

PRE_BUILD_DIR = wd / "install"

def main():
    if PRE_BUILD_DIR.exists():
        shutil.rmtree(PRE_BUILD_DIR)
    PRE_BUILD_DIR.mkdir()
    for p0, p1 in [((wd / p0), PRE_BUILD_DIR / p1) for p0, p1 in paths]:
        if p0.is_dir():
            shutil.copytree(p0, p1)
            for fp in p1.rglob("*"):
                if not fp.is_dir() and fp.name.startswith("~$"):
                    try:
                        os.remove(fp)
                    except Exception as e:
                        print(e)  

        else:
            if not p1.parent.exists():
                p1.parent.mkdir(parents=True)
            shutil.copy2(p0, p1)




if __name__ == "__main__":
    main()