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

if __name__ == "__main__":
    main()