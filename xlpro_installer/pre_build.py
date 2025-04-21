# Builds all code and moves it to the assets dir


from pathlib import Path
import subprocess
import shutil

# import xlpro.version

wd = Path(__file__).parent

import xlpro

paths = [
    "startfiles",
    f"../xlpro_module/dist/xlpro-{xlpro.__version__}-py3-none-any.whl",
    f"../xlpro_module/config.toml",
    f"../xlpro_addin/xlpro.xlam",
    f"../xlpro_cli/dist/xlpro-cli.exe",
]

out_dir = wd / "assets"
out_dir.mkdir(exist_ok=True)

# if not list(x for x in out_dir.glob("*")):
#     raise FileExistsError("Please clear out the assets directory")

for p in [(wd / p) for p in paths]:
    if p.is_dir():
        shutil.copytree(p, out_dir / p.name)
    else:
        shutil.copy2(p, out_dir / p.name)

pass
# pass
# res = subprocess.run(
#     [
#         ".venv\Scripts\pyinstaller.exe",
#         "main.spec",
#     ],
#     cwd=wd,
#     capture_output=True,
#     # check=True,
# )

# print(res.stderr)
# print(res.stdout)
