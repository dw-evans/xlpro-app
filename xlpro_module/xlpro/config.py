from __future__ import annotations
import toml
from pathlib import Path
from dataclasses import dataclass, field
import sys
import os


XLPRO_WD = (Path(os.environ["USERPROFILE"]) / ".xlpro").resolve()
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


if __name__ == "__main__":
    config = load()

pass