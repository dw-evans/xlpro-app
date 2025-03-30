from __future__ import annotations
import toml
from pathlib import Path
from dataclasses import dataclass, field
import sys


# def is_pyinstaller():
#     return hasattr(sys, '_MEIPASS')
# if is_pyinstaller():
#     XLPRO_WD = Path(r"C:\Users\Daniel Evans\.xlpro").parent.resolve()
# else:
#     XLPRO_WD = Path().resolve()
print("warning, fetching a specific config.toml path!")
XLPRO_WD = Path(r"C:\Users\Daniel Evans\.xlpro").resolve()
config_path = XLPRO_WD / "config.toml"

def load() -> Configuration:
    with open(config_path, "r") as f:
        kwargs = toml.load(f)
    settings = Configuration(**kwargs)
    return settings


@dataclass
class Configuration:
    # progid: str
    clsid: str
    python_path: str
    run_server_path:str
    debug_ip:str
    debug_port:int

    logging_level:str = field(default="DEBUG")
    logging_path:str = field(default="./xlpro.log")
    
    xlpro_lock_path:str = field(default="./xlpro.lock")

    xlpro_directory:str = field(default="./.xlpro")
    xlpro_functions_stem:str = field(default="functions")
    xlpro_subroutines_stem:str = field(default="subroutines")

    max_worker_threads:int = field(default=6)


if __name__ == "__main__":
    config = load()

pass