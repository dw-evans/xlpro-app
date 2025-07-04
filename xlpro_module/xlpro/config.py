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

import os

@dataclass
class Configuration:
    VSCODE_PATH:str = field(default="code.exe")
    XLPRO_CLI_PATH:str = field(default="xlpro-cli.exe")
    XLPRO_SERVER_PATH:str = field(default="xlpro-server.exe")
    LOGGING_LEVEL:str = field(default="INFO")
    MAX_WORKERS:int = field(default=12)
    UNDO_STACK_DEPTH:int = field(default=32)
    MULTI_SERVER_EXPERIEMENT:bool = field(default=False)
    xlpro_functions_stem:str = field(default="functions")
    xlpro_subroutines_stem:str = field(default="subroutines")

    def to_dict(self):
        attrs = (
            "xlpro_functions_stem",
            "xlpro_subroutines_stem",
            "MAX_WORKERS",
            "LOGGING_LEVEL",
            "VSCODE_PATH",
            "XLPRO_CLI_PATH",
            "UNDO_STACK_DEPTH",
            "MULTI_SERVER_EXPERIEMENT",
        )
        ret = {attr: getattr(self, attr) for attr in attrs}
        return ret


if __name__ == "__main__":
    config = load()

pass