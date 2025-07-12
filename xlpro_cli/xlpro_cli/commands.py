
import os
import ctypes

def enable_ansi_escape_codes_in_console():
    # Enable ANSI escape codes (24-bit color)
    kernel32 = ctypes.windll.kernel32
    handle = kernel32.GetStdHandle(-11)
    mode = ctypes.c_uint32()
    kernel32.GetConsoleMode(handle, ctypes.byref(mode))
    kernel32.SetConsoleMode(handle, mode.value | 0x0004)

enable_ansi_escape_codes_in_console()

from pathlib import Path
import subprocess
from . import utils

import sys
import argparse

from rich import print
import sys
import os

# change the os working directory... Not sure why.....
os.chdir(Path(__file__).parent.parent.parent)

@utils.try_except_press_enter_to_exit_wrapper
@utils.traceback_log_raise
def handle_start_server(args):
    """Starts the server"""
    workbook_path = Path(args.workbook)

    if not workbook_path.exists():
        raise FileNotFoundError(f"The provided workbook path does not exist {workbook_path}. Save the file and try again.")
    
    elif not utils.is_existing_xlpro_workbook_folder(workbook_path):
        raise Exception("workbook must be initialized for xlpro before launching")
    
    utils.start_venv_xlpro_server_for_workbook(workbook_path)
    


@utils.try_except_press_enter_to_exit_wrapper
@utils.traceback_log_raise
def handle_init(args):
    """initializes or re-initializes the workbook"""

    workbook_path = Path(args.workbook)

    if not workbook_path.exists():
        raise FileNotFoundError(f"The provided workbook path does not exist {workbook_path}. Save the file and try again.")
    
    utils.dlg_xlpro_initialize_workbook(workbook_path)
    utils.press_enter_or_timeout_exit(timeout=30.0)


@utils.try_except_press_enter_to_exit_wrapper
@utils.traceback_log_raise
def handle_clear_local_mapping_for_workbook(args):
    """Removes the venv mapping for the workbook"""

    workbook_path = Path(args.workbook)

    if not workbook_path.exists():
        raise FileNotFoundError(f"The provided workbook path does not exist {workbook_path}. Save the file and try again.")
    
    utils.remove_venv_mapping_for_workbook(workbook_path, dialogue=True)

    utils.press_enter_or_timeout_exit(timeout=30.0)


@utils.try_except_press_enter_to_exit_wrapper
@utils.traceback_log_raise
def _handle_get_guid(args):
    workbook_path = Path(args.workbook)
    
    # py_interpreter_root_dir = utils.get_valid_venv_root_path_used_for_workbook_from_map(workbook_path)
    # py_interpreter_path = utils.get_python_exe_from_xlpro_root_venv_path(py_interpreter_root_dir)
    guid = utils.get_interpreter_guid(workbook_path)
    sys.stderr.write(str(guid))
    pass

@utils.try_except_press_enter_to_exit_wrapper
@utils.traceback_log_raise
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
        raise e
    print(f"Requirements written successfully for {workbook_path}")
    # utils.press_enter_to_exit()
    utils.press_enter_or_timeout_exit(timeout=30.0)

@utils.try_except_press_enter_to_exit_wrapper
@utils.traceback_log_raise
def handle_clear_venv_data(args):
    utils.check_envs_folder_size_prompt_delete()
    # utils.press_enter_to_exit()
    utils.press_enter_or_timeout_exit(timeout=5.0)

@utils.try_except_press_enter_to_exit_wrapper
@utils.traceback_log_raise
def handle_clear_tmp_data(args):
    utils.check_folder_size_prompt_delete(utils.XLPRO_TMP_FOLDER_PATH)
    utils.press_enter_or_timeout_exit(timeout=5.0)

@utils.try_except_press_enter_to_exit_wrapper
@utils.traceback_log_raise
def handle_clear_uv_cache_data(args):
    utils.check_folder_size_prompt_delete(Path(os.environ["UV_CACHE_DIR"]))
    utils.press_enter_or_timeout_exit(timeout=5.0)

@utils.try_except_press_enter_to_exit_wrapper
@utils.traceback_log_raise
def handle_clear_uv_pythons_data(args):
    utils.check_folder_size_prompt_delete(Path(os.environ["UV_PYTHON_INSTALL_DIR"]))
    utils.press_enter_or_timeout_exit(timeout=5.0)


XLPRO_INSTALL_DIR = (Path(os.environ["USERPROFILE"]) / ".xlpro").resolve()
XLPRO_BIN_DIR =  XLPRO_INSTALL_DIR / "bin"

def _configure_env():
    os.environ["PATH"] = f"{XLPRO_BIN_DIR};" + os.environ["PATH"]
    uv_cache_dir = XLPRO_INSTALL_DIR / "uv/cache"
    uv_cache_dir.mkdir(exist_ok=True, parents=True)
    os.environ["UV_CACHE_DIR"] = str(uv_cache_dir.resolve())
    uv_python_dir = XLPRO_INSTALL_DIR / "uv/python"
    uv_python_dir.mkdir(exist_ok=True)
    os.environ["UV_PYTHON_INSTALL_DIR"] = str(uv_python_dir.resolve())

def _load_uv_help() -> str:
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


# def get_cli_header_path():
#     if getattr(sys, 'frozen', False):
#         print("running in meipass")
#         # Running in PyInstaller bundle
#         base_path = Path(sys._MEIPASS)

#         import pkgutil
#         header = pkgutil.get_data(__name__, "cli-header.txt").decode("utf-8")

#     else:
#         print("running in default mode")
#         # Running in development
#         base_path = Path(__file__).parent.parent

#     ret = base_path / "cli-header.txt"
#     print(ret, ret.exists())
#     return ret

import tempfile

def main():

    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
    # with open(get_cli_header_path(), "r", encoding="utf-8") as f:
    #     print(f.read())
    try:
        if getattr(sys, 'frozen', False):
            base_path = Path(sys._MEIPASS)
            fp = base_path / "cli-header.txt"
            # print(fp, fp.exists())
            with open(fp, "r", encoding="utf-8") as f:
                print(f.read())
        
        else:
            # print("running in default mode")
            # Running in development
            base_path = Path(__file__).parent.parent
            with open(base_path / "cli-header.txt", "r", encoding="utf-8") as f:
                print(f.read())
    except Exception as e:
        print("Could not load cli-header.txt")

    _configure_env()

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
    parser_get_guid = subparsers.add_parser("guid", help="get the guid for a workbook (if the process is active)")
    parser_write_reqs = subparsers.add_parser("write-reqs", help="write the requirements to the sever location")
    parser_clear_wb_venv = subparsers.add_parser("clear-venv-link", help="Remove the venv link for the workbook")
    
    parser_delete_venv_dir = subparsers.add_parser("clear-venvs", help="Remove all virtual environment data")
    parser_delete_tmp_dir = subparsers.add_parser("clear-tmp", help="Remove all temporary data")
    parser_delete_uv_cache = subparsers.add_parser("clear-uv-cache", help="Remove cached uv data")
    parser_delete_uv_pythons = subparsers.add_parser("clear-uv-pythons", help="Remove cached python installations")

    parser_start.add_argument("workbook", type=str, help="workbook to start xlpro server for")
    parser_init.add_argument("workbook", type=str, help="workbook to initialize xlpro for (writes adjacent folder structure)")
    parser_get_guid.add_argument("workbook", type=str, help="workbook to get running server for")
    parser_write_reqs.add_argument("workbook", type=str, help="workbook to write requirements for.")
    parser_clear_wb_venv.add_argument("workbook", type=str, help="workbook to write requirements for.")
    parser_write_reqs.add_argument("--force", action="store_true", help="force the update", required=False)

    parser_start.set_defaults(func=handle_start_server)
    parser_init.set_defaults(func=handle_init)
    parser_get_guid.set_defaults(func=_handle_get_guid)
    parser_write_reqs.set_defaults(func=handle_write_requirements)
    parser_delete_venv_dir.set_defaults(func=handle_clear_venv_data)
    parser_delete_tmp_dir.set_defaults(func=handle_clear_tmp_data)
    parser_delete_uv_cache.set_defaults(func=handle_clear_uv_cache_data)
    parser_delete_uv_pythons.set_defaults(func=handle_clear_uv_pythons_data)
    parser_clear_wb_venv.set_defaults(func=handle_clear_local_mapping_for_workbook)


    args = parser.parse_args()
    args.func(args)



if __name__ == "__main__":
    try:
        main()
    except SystemExit as e:
        raise
    except KeyboardInterrupt as e:
        raise
    except Exception as e:
        print(f"Exception encountered: {e}")
        input("Press enter to Close")
    