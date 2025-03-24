# from xlpro_cli import commands
# from xlpro_cli import utils
# from pathlib import Path
import sys
sys.path.append(r"C:\Users\Daniel Evans\projects\xlpro")
# sys.path.append(r"C:\Users\Daniel Evans\projects\xlpro\xlpro_cli")
from xlpro_cli.xlpro_cli import commands
from xlpro_cli.xlpro_cli import utils
from pathlib import Path

def main():
    fp = Path() / r"C:\Users\Daniel Evans\projects\xlpro\xlpro_testing\test3-initialize-from-existing-system\book1.xlsx"
    utils.start_venv_xlpro_server_for_workbook(fp)


def test_install_default_requirements():
    utils.install_default_requirements(Path() / r"C:\Users\Daniel Evans\projects\xlpro\xlpro_install\envs\30baec6d-aa3c-4f55-983a-3cd8f1ae65e9_3.13.2\.venv\Scripts\python.exe")
    pass


if __name__ == "__main__":
    test_install_default_requirements()
