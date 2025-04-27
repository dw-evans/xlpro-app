
from pathlib import Path
import subprocess
from . import utils
import sys
import argparse

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
    # parser_uv = subparsers.add_parser("uv", help=load_uv_help())

    parser_start.add_argument("workbook", type=str, help="workbook to start xlpro server for")

    parser_init.add_argument("workbook", type=str, help="workbook to initialize xlpro for (writes adjacent folder structure)")
    parser_uninit.add_argument("workbook", type=str, help="workbook to uninitialize xlpro for (removes adjacent folder structure)")
    parser_get_guid.add_argument("workbook", type=str, help="workbook to uninitialize xlpro for (removes adjacent folder structure)")

    parser_start.set_defaults(func=handle_start_server)
    parser_init.set_defaults(func=handle_init)
    parser_uninit.set_defaults(func=handle_uninit)
    parser_get_guid.set_defaults(func=handle_get_guid)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()