import sys
from pathlib import Path

wd = Path(__file__).parent

EXTRA_PATHS = [
    wd / "../../example_utils.py",

]

sys.path += EXTRA_PATHS
