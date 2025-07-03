# Builds all code and moves it to the assets dir


from pathlib import Path
import subprocess
import shutil

# import xlpro.version

wd = Path(__file__).parent

import xlpro

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
        f"../xlpro_module/config.toml",
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
        f"LICENSE",
    ),
]

import os
import shutil

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