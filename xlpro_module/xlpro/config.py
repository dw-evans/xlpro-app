from __future__ import annotations
import toml
from pathlib import Path
from dataclasses import dataclass, field
import sys
import os


def is_pyinstaller():
    return hasattr(sys, '_MEIPASS')

if is_pyinstaller():
    XLPRO_WD = Path(sys.executable).parent.resolve()
    # XLPRO_WD = Path(os.environ.get("USERPROFILE")) / ".xlpro"
else:
    # development path manually set to the installation directory
    XLPRO_WD = Path(os.environ.get("USERPROFILE")) / ".xlpro"

config_path = XLPRO_WD / "config.toml"

def load() -> Configuration:
    with open(config_path, "r") as f:
        kwargs = toml.load(f)
    settings = Configuration(**kwargs)
    return settings


@dataclass
class Configuration:
    xlpro_functions_stem:str
    xlpro_subroutines_stem:str
    max_worker_threads:int
    logging_level:str
    vscode_path:str
    xlpro_cli_path:str

    def to_dict(self):
        attrs = (
            "xlpro_functions_stem",
            "xlpro_subroutines_stem",
            "max_worker_threads",
            "logging_level",
            "vscode_path",
            "xlpro_cli_path",
        )
        ret = {attr: getattr(self, attr) for attr in attrs}
        return ret


if __name__ == "__main__":
    config = load()

pass