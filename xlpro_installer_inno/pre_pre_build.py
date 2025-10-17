# Builds all code and moves it to the assets dir


from pathlib import Path
import subprocess
import shutil
import os

import xlpro
import shutil
import logging

wd = Path(__file__).parent

paths = [
    # (
    #     "startfiles",
    #     "startfiles",
    # ),
    (
        f"../xlpro/dist/xlpro-{xlpro.__version__}-py3-none-any.whl",
        f"src/xlpro-{xlpro.__version__}-py3-none-any.whl",
    ),
    (f"../xlpro/config_production.toml", "config.toml"),
    (
        f"../xlpro_addin/dist/xlpro.xlam",
        f"src/xlpro.xlam",
    ),
    (
        f"../xlpro_cli/dist/xlpro-cli.exe",
        "xlpro-cli.exe",
    ),
    # (
    #     f"../xlpro_examples",
    #     f"examples",
    # ),
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
        f"../licenses/third_party_licenses.txt",
        f"licenses/third_party_licenses.txt",
    ),
    (
        f"../xlpro_cli_shell_wrapper/xlpro-server.exe",
        f"xlpro-server.exe",
    ),
    # Copy recommended python version to correct location
    (
        f".python-version-recommended",
        f".python-version-recommended",
    ),
    (
        # https://github.com/astral-sh/uv/releases/download/0.6.10/uv-x86_64-pc-windows-msvc.zip
        f"uv/uv.exe",
        f"bin/uv.exe",
    ),
]

PRE_BUILD_DIR = wd / "build"


def collect_examples():
    r = []
    for fp in (wd / Path("../xlpro_examples")).rglob("*"):
        if fp.is_dir():
            continue
        if any([y in ["__pycache__", ".vscode", "Modified", "Original"] for y in fp.parts]):
            continue
        if fp.name.startswith("~$"):
            continue
        r.append(fp.relative_to(wd))
        pass
    return r


def collect_startfiles():
    return [x.relative_to(wd) for x in (wd / "startfiles").rglob("*")]


paths += [(str(fp), str(Path("examples") / fp)) for fp in collect_examples()]
paths += [(str(fp), str(fp)) for fp in collect_startfiles()]


def main():
    if PRE_BUILD_DIR.exists():
        shutil.rmtree(PRE_BUILD_DIR)
    PRE_BUILD_DIR.mkdir()
    for p0, p1 in [((wd / p0), PRE_BUILD_DIR / p1) for p0, p1 in paths]:
        # if p0.is_dir():
        #     shutil.copytree(p0, p1)
        #     for fp in p1.rglob("*"):
        #         if not fp.is_dir() and fp.name.startswith("~$"):
        #             try:
        #                 os.remove(fp)
        #             except Exception as e:
        #                 print(e)

        if not p1.parent.exists():
            p1.parent.mkdir(parents=True)
        shutil.copy2(p0, p1)


if __name__ == "__main__":
    main()
    pass
