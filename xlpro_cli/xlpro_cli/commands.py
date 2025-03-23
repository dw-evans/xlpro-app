
from pathlib import Path
import subprocess
from xlpro_cli.xlpro_cli import utils
import sys
import argparse


# parser = argparse.ArgumentParser(prog='xlpro-cli')

def start_server():
    parser = argparse.ArgumentParser("Run the xlpro server")
    # parser.add_argument("--debug", action="store_true", help="enable debugpy debugging")
    parser.add_argument("--workbook", type=str, required=True, help="Workbook to initialize xlpro for")
    args = parser.parse_args()
    workbook_path = Path(args.workbook)
    
    if not workbook_path.exists():
        raise FileNotFoundError
    elif utils.is_existing_xlpro_workbook_folder(workbook_path):
        raise Exception("workbook must be initialized for xlpro before launching")
    
    utils.start_venv_xlpro_server_for_workbook(workbook_path)
    pass

    
if __name__ == "__main__":
    start_server()