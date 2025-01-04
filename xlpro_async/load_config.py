from __future__ import annotations
import toml
from pathlib import Path
from dataclasses import dataclass

wd = Path(__file__).parent
config_path = wd / "config.toml"

def load_config() -> Settings:
    with open(config_path, "r") as f:
        kwargs = toml.load(f)
    settings = Settings(**kwargs)
    return settings


@dataclass
class Settings:
    progid: str
    clsid: str
    python_path: str
    logging_level:str
    logging_path:str
    
    run_server_path:str



if __name__ == "__main__":
    config = load_config()

pass