
from pathlib import Path
import subprocess
from . import utils
import sys
import argparse

def handle_start_server(args):
    workbook_path = Path(args.workbook)

    if not workbook_path.exists():
        raise FileNotFoundError
    
    elif not utils.is_existing_xlpro_workbook_folder(workbook_path):
        raise Exception("workbook must be initialized for xlpro before launching")
    
    utils.start_venv_xlpro_server_for_workbook(workbook_path)
    pass


def handle_init(args):
    """initializes or re-initializes the workbook"""
    workbook_path = Path(args.workbook)

    if not workbook_path.exists():
        raise FileNotFoundError
    
    utils.dlg_xlpro_initialize_workbook(workbook_path)


def handle_uninit(args):
    raise NotImplementedError


def main():
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

    parser_start.add_argument("workbook", type=str, help="workbook to start xlpro server for")
    parser_init.add_argument("workbook", type=str, help="workbook to initialize xlpro for (writes adjacent folder structure)")
    parser_uninit.add_argument("workbook", type=str, help="workbook to uninitialize xlpro for (removes adjacent folder structure)")

    parser_start.set_defaults(func=handle_start_server)
    parser_init.set_defaults(func=handle_init)
    parser_uninit.set_defaults(func=handle_uninit)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    handle_start_server()