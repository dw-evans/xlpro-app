
from pathlib import Path
import subprocess
from . import utils
import sys
import argparse

from rich.console import Console
from rich import print
from rich.style import Style

import os
os.chdir(Path(__file__).parent.parent.parent)

def handle_start_server(args):
    """Starts the server"""
    workbook_path = Path(args.workbook)

    if not workbook_path.exists():
        raise FileNotFoundError
    
    elif not utils.is_existing_xlpro_workbook_folder(workbook_path):
        raise Exception("workbook must be initialized for xlpro before launching")
    
    utils.start_venv_xlpro_server_for_workbook(workbook_path)
    pass


def handle_register_workbook_on_server(args):
    ...


def handle_init(args):
    """initializes or re-initializes the workbook"""
    workbook_path = Path(args.workbook)

    if not workbook_path.exists():
        raise FileNotFoundError
    
    utils.dlg_xlpro_initialize_workbook(workbook_path)


def handle_uninit(args):
    raise NotImplementedError

import pythoncom
def handle_get_guid(args):
    workbook_path = Path(args.workbook)
    # py_interpreter_root_dir = utils.get_valid_venv_root_path_used_for_workbook_from_map(workbook_path)
    # py_interpreter_path = utils.get_python_exe_from_xlpro_root_venv_path(py_interpreter_root_dir)
    guid = utils.get_interpreter_guid(workbook_path)
    sys.stderr.write(str(guid))
    pass

def handle_write_requirements(args):
    """ Writes requirements txt, .python-version, .xlpro-version - pending...
    """
    workbook_path = Path(args.workbook)
    if not args.force:
        if utils.prompt_yes_no_input(f"You are about to write the requirements for {workbook_path}, are you sure?", "yes") == "no":
            utils.print_info("exiting")
    try:
        utils.print_info("Writing requirements...")
        utils.write_reqs_for_workbook(workbook_path)
        utils.print_info("Requirements written successfully")

        utils.print_info("Writing python-version...")
        utils.write_python_version_file_for_workbook(workbook_path)
        utils.print_info(".python-version written successfully")
        
    except Exception as e:
        print(f"Exception occured when attempting to write workbook requirements data: {e}")
        return
    print(f"Requirements written successfully for {workbook_path}")


XLPRO_INSTALL_DIR = (Path(os.environ["USERPROFILE"]) / ".xlpro").resolve()
XLPRO_BIN_DIR =  XLPRO_INSTALL_DIR / "bin"

def configure_env():
    os.environ["PATH"] = f"{XLPRO_BIN_DIR};" + os.environ["PATH"]

def load_uv_help() -> str:
    result = subprocess.run(
        [
            "uv",
            "-h"
        ],
        check=True,
        capture_output=True,
        text=True,
    )
    ret = result.stdout
    return ret


def main():
    configure_env()
    load_uv_help()

    parser = argparse.ArgumentParser(
        prog="xlpro-cli",
        description="xlpro command-line utility",
    )
    subparsers = parser.add_subparsers(
        dest="command",
        required=True,
        help=""
    )

    parser_start = subparsers.add_parser("start", help="run the xlpro server")
    parser_init = subparsers.add_parser("init", help="initialize a workbook for xlpro")
    parser_uninit = subparsers.add_parser("uninit", help="uninitialize a workbook for xlpro")
    parser_get_guid = subparsers.add_parser("guid", help="get the guid for a workbook (if the process is active?)")
    parser_write_reqs = subparsers.add_parser("write-reqs", help="write the requirements to the sever location")

    parser_start.add_argument("workbook", type=str, help="workbook to start xlpro server for")
    parser_init.add_argument("workbook", type=str, help="workbook to initialize xlpro for (writes adjacent folder structure)")
    parser_uninit.add_argument("workbook", type=str, help="workbook to uninitialize xlpro for (removes adjacent folder structure)")
    parser_get_guid.add_argument("workbook", type=str, help="workbook to get running serverfor (removes adjacent folder structure)")
    
    parser_write_reqs.add_argument("workbook", type=str, help="workbook to write requirements for.")
    parser_write_reqs.add_argument("--force", action="store_true", help="force the update", required=False)

    parser_start.set_defaults(func=handle_start_server)
    parser_init.set_defaults(func=handle_init)
    parser_uninit.set_defaults(func=handle_uninit)
    parser_get_guid.set_defaults(func=handle_get_guid)
    parser_write_reqs.set_defaults(func=handle_write_requirements)

    args = parser.parse_args()
    args.func(args)

import sys
import time

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Exception encountered: {e}")
    finally:
        input("press enter to continue")
    time.sleep(5)
    