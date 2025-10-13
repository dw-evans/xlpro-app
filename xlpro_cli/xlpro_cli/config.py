from __future__ import annotations

# import tomllib
import tomli
from pathlib import Path
from dataclasses import dataclass, field, fields
import sys
import os
import logging
from typing import Any, get_type_hints

logger = logging.getLogger(__name__)


def is_pyinstaller():
    return hasattr(sys, "_MEIPASS")


if is_pyinstaller():
    XLPRO_WD = Path(sys.executable).parent.resolve()
    # XLPRO_WD = Path(os.environ.get("USERPROFILE")) / ".xlpro"
else:
    # development path manually set to the installation directory
    XLPRO_WD = Path(os.environ.get("USERPROFILE")) / ".xlpro"

config_path = XLPRO_WD / "config.toml"


def load() -> Configuration:
    if not config_path.exists():
        print(f"WARNING, CONFIG PATH NOT FOUND AT {config_path}")
        kwargs = {}
    else:
        err = Configuration.validate_file(config_path)
        if err:
            logger.error(f"TOML Decode error: {err}")
            raise err

        with open(config_path, "rb") as f:
            kwargs = tomli.load(f)
        err = Configuration.validate_kwargs(kwargs)
        if err:
            logger.error(f"Configuration values invalid, see output below")
            for e in err:
                logger.error(f"{e}")
            raise Exception("Configuration file invalid")

    settings = Configuration(**kwargs)
    return settings


@dataclass
class Configuration:
    VSCODE_PATH: str = field(default="code.exe")
    XLPRO_CLI_PATH: str = field(default="xlpro-cli.exe")
    XLPRO_SERVER_PATH: str = field(default="xlpro-server.exe")
    LOGGING_LEVEL: str = field(default="INFO")
    MAX_WORKERS: int = field(default=12)
    UNDO_STACK_DEPTH: int = field(default=32)
    MULTI_SERVER_EXPERIEMENT: bool = field(default=False)
    xlpro_functions_stem: str = field(default="functions")
    xlpro_subroutines_stem: str = field(default="subroutines")

    @staticmethod
    def validate_file(fp: Path):
        e = None
        try:
            with open(fp, "rb") as f:
                # x = tomllib.load(f)
                x = tomli.load(f)
            return
        except tomli.TOMLDecodeError as e:
            return e

    @staticmethod
    def validate_kwargs(kwargs: dict[str, Any]) -> list[str]:
        errs = []
        type_hints = get_type_hints(Configuration)  # ensures resolved types (e.g., str, int)

        for k, v in kwargs.items():
            if k not in type_hints:
                errs.append(f"Invalid config key: '{k}'")
                continue

            expected_type = type_hints[k]
            if not isinstance(v, expected_type):
                errs.append(f"Invalid type for '{k}': expected {expected_type.__name__}, got {type(v).__name__}")

        return errs  # always a list, empty if no errors

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
