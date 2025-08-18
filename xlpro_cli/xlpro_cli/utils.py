from rich.console import Console
from rich import print
from rich.style import Style
import readchar
import sys
import subprocess
import time
from pathlib import Path
# import regex as re
import json
import logging
import os
import uuid
import shutil
import socket
import winreg
import psutil
from pathlib import Path
import sys
import textwrap
import re
import stat
from . import config
from . import file_lock
import datetime


# import io
# sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
# sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
from functools import wraps
import traceback

# DEVELOPMENT_INSTALL = True
DEVELOPMENT_INSTALL = False

logging.basicConfig(
    stream=sys.stdout,
    level=logging.DEBUG,          # The log level (DEBUG, INFO, WARNING, etc.)
    format='%(asctime)s - %(levelname)s - %(message)s',  # The format of log messages
    datefmt='%Y-%m-%d %H:%M:%S'    # The format of the date in log messages
)

logger = logging.getLogger(__name__)

def press_enter_to_exit():
    input("Press enter to exit")

def try_except_press_enter_to_exit_wrapper(func):
    @wraps(func)
    def inner(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except SystemExit as e:
            raise
        except KeyboardInterrupt as e:
            raise
        except Exception as e:
            print(f"Error occurred. {e}")
            traceback.print_exc()
            # utils.press_enter_or_timeout_exit(timeout=)
            press_enter_to_exit()
    return inner




def is_pyinstaller():
    return hasattr(sys, '_MEIPASS')

if is_pyinstaller():
    XLPRO_WD = Path(sys.executable).parent.resolve()
    # XLPRO_WD = Path(os.environ.get("USERPROFILE")) / ".xlpro"
else:
    XLPRO_WD = Path(os.environ.get("USERPROFILE")) / ".xlpro"

# print("cwd is " + os.getcwd())
# XLPRO_ROOT_PATH = XLPRO_WD / "xlpro_install"
XLPRO_ROOT_PATH = XLPRO_WD
XLPRO_ENVS_DIR = XLPRO_ROOT_PATH / 'envs'
XLPRO_VENV_WORKBOOKS_MAP_JSON_FP = XLPRO_ENVS_DIR / "venv-mappings.json"
XLPRO_SRC_DIR = XLPRO_WD / "src"
XLPRO_TMP_FOLDER_PATH = XLPRO_ROOT_PATH / "tmp"


def sleep_then_close(timeout_sec=1.0):
    console.print(f"Closing in {timeout_sec}...")
    time.sleep(timeout_sec)
    sys.exit(0)


def press_enter_or_timeout_exit(timeout=5):
    import threading

    enter_event = threading.Event()

    def wait_for_input():
        input(f"Press Enter to exit (timeout in {timeout} sec)...")
        enter_event.set()

    t1 = threading.Thread(target=wait_for_input, daemon=True)
    t1.start()

    if enter_event.wait(timeout=timeout):
        print("Enter Pressed. Closing...")
        sys.exit(0)

    print(f"Timeout reached. Closing...")
    sys.exit(0)

    return
    # def force_exit_after_timeout():
    #     time.sleep(timeout)
    #     print(f"\nTimeout reached after {timeout} seconds. Exiting.")
    #     sys.exit()  # non-zero exit code for timeout

    # # Start the timeout thread
    # timeout_thread = threading.Thread(target=force_exit_after_timeout, daemon=True)
    # timeout_thread.start()

    # # Prompt user on main thread
    # input(f"Press Enter to exit (timeout in {timeout} sec)...\n")
    # print("Closing...")
    # timeout_thread.join()
    # sys.exit()

def get_terminal_width() -> int:
    return shutil.get_terminal_size().columns


def clear_text(text:str) -> None:
    terminal_width = shutil.get_terminal_size().columns
    lines = text.split("\n")
    for line in lines:
        line_count_of_line = (len(line) // terminal_width) + 1
        for _ in range(line_count_of_line):
            sys.stdout.write("\033[A\r" + f" " * terminal_width + "\r")  # Overwrite the line with spaces
    sys.stdout.flush()


def clear_line(count=1):
    for _ in range(count):
        sys.stdout.write("\033[A\r" + f" " * 200 + "\r")  # Overwrite the line with spaces
    sys.stdout.flush()

pass

console = Console(highlight=False)

style_prompt = Style.parse("#526cfe")
style_prompt_boldface = style_prompt + Style.parse("bold")

style_generic_option = Style.parse("#cccccc")
style_selected_option = style_generic_option + Style.parse("bold") + Style.parse("reverse")

style_plain = Style.parse("#cccccc")
style_plain_boldface = style_plain + Style.parse("bold")

style_success = Style.parse("#526cfe")
style_success_boldface = style_success + Style.parse("bold")

style_error = Style.parse("#e5342f")
style_error_boldface = style_error + Style.parse("bold")

style_warning = Style(color="#eeba56")
style_warning_boldface = style_warning + Style.parse("bold")


def initialize_xlpro_install_directory():
    try:
        XLPRO_ROOT_PATH.mkdir(exist_ok=True)
    except Exception as e:
        raise e

    if not XLPRO_VENV_WORKBOOKS_MAP_JSON_FP.exists():
        with open(XLPRO_VENV_WORKBOOKS_MAP_JSON_FP, "w") as f:
            f.write(json.dumps({}))

    (XLPRO_ROOT_PATH / "envs").mkdir(exist_ok=True)


def create_uuid_str():
    return str(uuid.uuid4())


def get_xlpro_python_interpreters() -> list[Path]:
    xlpro_venv_interepreters = [str(v.resolve()) for x in XLPRO_ENVS_DIR.glob("*") if x.is_dir() and (v:=(x / ".venv/scripts/python.exe")).exists()]
    xlpro_venv_interepreters.sort()
    return xlpro_venv_interepreters


def get_uv_python_interpreters() -> list[Path]:
    # result = subprocess.run("uv python dir", shell=True, capture_output=True, text=True, check=True)
    result = subprocess.run(["uv", "python", "dir"], capture_output=True, text=True, check=True)
    uv_py_dir = Path(result.stdout.split("\n")[0])
    uv_py_exes = [x for x in uv_py_dir.glob("*/python.exe")]
    uv_py_exes.sort()
    return uv_py_exes


def get_global_python_interpreters():
    # result = subprocess.run("where.exe python", shell=True, capture_output=True, text=True, check=True)
    result = subprocess.run(["where.exe", "python"], capture_output=True, text=True, check=True)
    ret = result.stdout.split("\n")[:-1]
    ret.sort()
    return ret


def uv_download_python_version(version_str:str) -> Path:
    print_info("Downloading Python interpreter from uv...")
    result = subprocess.run(
        [
            "uv", 
            "python", 
            "install",
            version_str,
            "--verbose"
        ],
        capture_output=True, 
        text=True, 
        check=True,
    )
    print_info("Download complete.")


    def uninstall():
        result = subprocess.run(
            [
                "uv", 
                "python", 
                "uninstall",
                version_str,
            ],
            capture_output=True, 
            text=True, 
            check=True,
        )

    uv_interpreter_path_list = get_uv_python_interpreters()

    found_path_list = [x for x in uv_interpreter_path_list if re.search(version_str.replace(r".", r"\."), str(x.parent))]

    if len(found_path_list) > 1:
        raise Exception("multiple files found")
    if not found_path_list:
        raise FileNotFoundError("could not find the just-installed uv interpreter.")
    
    return found_path_list[0]
    

def is_interpreter_valid_venv_and_exists(py_interpreter_path:Path) -> tuple[bool, Exception|None]:
    if not py_interpreter_path.name == "python.exe":
        return False, Exception("provided path should be path/to/python.exe executable")
    if not py_interpreter_path.exists():
        return False, FileNotFoundError("file does not exist")
    # ./.venv/pyvenv.cfg will exist for virtual environments
    if (py_interpreter_path.parent.parent / "pyvenv.cfg").exists():
         return True, None
    return False, Exception("other error")


def get_t2_version_str(s:str) -> str:
    return re.search(r'(\d\.\d+)', s).group(1)


def compare_py_version_t2(py_version_1:str, py_version_2:str) -> bool:
    return get_t2_version_str(py_version_1) == get_t2_version_str(py_version_2)
    

def get_py_exe_version(py_interpreter_path:Path) -> str:
    str_py_interpreter_path = str(py_interpreter_path)
    result = subprocess.run(
        [
            str_py_interpreter_path, 
            "--version",
        ], 
        shell=True, 
        capture_output=True, 
        text=True, 
        check=True,
    )
    version_match = re.search(r"Python (\d+\.\d+\.\d+)", result.stdout.split("\n")[0])
    return version_match.group(1)


def generate_xlpro_venv_dir_name(py_interpreter_path:Path):
    chosen_py_version = get_py_exe_version(py_interpreter_path)
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d")
    # return f"{timestamp}_{create_uuid_str()}_{chosen_py_version}"
    return f"{timestamp}_Py-{chosen_py_version}"


def create_xlpro_venv_from_interpreter_and_get_root_path(py_interpreter_path:Path) -> Path:
    """Creates a venv in the xlpro env folder and returns the path to the .venv parent directory"""
    fp_exists = True
    i = 0
    while fp_exists:
        xlpro_venv_path:Path = XLPRO_ENVS_DIR / (generate_xlpro_venv_dir_name(py_interpreter_path) + f"{('_' + str(i)) if i > 0 else ''}") /'.venv'
        fp_exists = xlpro_venv_path.exists()
        i += 1
    xlpro_venv_path.parent.mkdir(exist_ok=True, parents=True)

    result = subprocess.run(
        args=[
            "uv",
            "venv",
            "--python",
            str(py_interpreter_path.resolve()),
            f"{xlpro_venv_path}",
        ],
        check=True,
    )
    # initialize_xlpro_venv_files(xlpro_venv_path)
    return xlpro_venv_path.resolve()


def create_symbolic_venv_from_existing_venv(existing_venv_root_path:Path) -> Path:
    """Creates a symlinked virtual environment in the xlpro directory using a junction. Must provide the root
    path to the existing virtual environment"""

    if not (v:=(existing_venv_root_path / "scripts/python.exe").resolve()).exists():
        logger.error(f"python.exe does not exist within {existing_venv_root_path}")
        raise Exception(f"provided environment is invalid (no {v} found)")

    py_interpreter_path = existing_venv_root_path / "scripts/python.exe"

    fp_exists = True
    i = 0
    while fp_exists:
        xlpro_venv_path:Path = XLPRO_ENVS_DIR / (generate_xlpro_venv_dir_name(py_interpreter_path) + f"{('_' + str(i)) if i > 0 else ''}") /'.venv'
        fp_exists = xlpro_venv_path.exists()
        i += 1
    xlpro_venv_path.parent.mkdir(exist_ok=True)

    if xlpro_venv_path.exists():
        raise Exception("venv junction link file already exists")

    result = subprocess.run(
        [
            "mklink",
            "/J",
            str(xlpro_venv_path.name),
            str(existing_venv_root_path.resolve()),
        ],
        cwd=str(xlpro_venv_path.resolve().parent),
        check=True,
        shell=True,
        capture_output=True,
    )

    return xlpro_venv_path


def initialize_xlpro_venv_files(venv_path:Path):
    """Temporary function to initalize the xlpro files for a virtual environment"""
    xlpro_sub_dir = venv_path.parent / ".xlpro"
    xlpro_sub_dir.mkdir()

    files_to_make = (
        xlpro_sub_dir / "xlpro.log",
        xlpro_sub_dir / "xlpro.cfg",
        xlpro_sub_dir / "xlpro.meta",
        xlpro_sub_dir / "pyproject.toml",
        xlpro_sub_dir / "requirements.txt",
        xlpro_sub_dir / ".python-version",
    )

    for fp in files_to_make:
        with open(fp, "w") as f:
            f.write("")


def get_xlpro_workbook_directory(workbook_path:Path) -> Path:
    return workbook_path.parent / f"{workbook_path.name}.xlpro"


def initialize_and_get_workspace_xlpro_dir(workbook_path:Path) -> Path:
    d = get_xlpro_workbook_directory(workbook_path)
    d.mkdir(exist_ok=True)

    startfile_dir = XLPRO_ROOT_PATH / "startfiles"
    startfile_contents = list(startfile_dir.glob("*"))

    for p in startfile_contents:
        if not (dst:=d / p.name).exists():
            shutil.copy2(p, dst)
        

    # funcs_path = d / f"functions.py"
    # subroutines_path = d / f"subroutines.py"

    # p = funcs_path
    # if not p.exists():
    #     with open(p, "w") as f:
    #         f.write(f"# > {p.resolve()}\n")
    #         f.write(f"# xlpro will automatically detect functions in this file as Excel UDFs.\n\n")

    # p = subroutines_path
    # if not p.exists():
    #     with open(p, "w") as f:
    #         f.write(f"# > {p.resolve()}\n")
    #         f.write(f"# xlpro will automatically detect functions in this file as Excel subroutines.\n\n")

        
    return d


def get_python_exe_from_xlpro_root_venv_path(root_venv_path:Path) -> Path:
    if (ret:=root_venv_path / "scripts/python.exe").exists():
        return ret
    elif (ret:=root_venv_path / ".venv/scripts/python.exe").exists():
        return ret
    raise FileNotFoundError(f"{root_venv_path} is not correctly formatted")


def validate_xlpro_workbook_folder(xlpro_folder:Path):
    """
    Requirements for an xlpro workbook folder:
        (optional) functions.py 
        (optional) subroutines.py 
        (optional) requirements.txt 
        (optional) .python-version
    """
    if not (v1:=(xlpro_folder / "functions.py")).exists():
        logger.warning(f"warning, recommended file is not present: {v1}")
    if not (v2:=(xlpro_folder / "subroutines.py")).exists():
        logger.warning(f"warning, recommended file is not present: {v2}")
    if not (v3:=(xlpro_folder / "requirements.txt")).exists():
        logger.warning(f"warning, recommended file is not present: {v3}")
    if not (v4:=(xlpro_folder / ".python-version")).exists():
        logger.warning(f"warning, recommended file is not present: {v4}")

    return True


def is_existing_xlpro_workbook_folder(workbook_path:Path) -> bool:
    """Check if the ...xlpro/ directory exists, validate and return true if so."""
    potential_dir = get_xlpro_workbook_directory(workbook_path)
    if potential_dir.exists():
        # validate_xlpro_workbook_folder(potential_dir)
        try:
            validate_xlpro_folder(potential_dir)
            return True
        except Exception as e:
            return False
        validate_xlpro_workbook_folder(potential_dir)
    return False


def reinitialize_workbook_for_xlpro(workbook_path:Path) -> Path:
    """Reinitialize the environment for this workbook. Fetches the cached venv used"""
    found_venv = get_valid_venv_root_path_used_for_workbook_from_map(workbook_path)

    # handle failed find 
    if not found_venv:
        logger.warning(f"the cached venv could not be found for {workbook_path}")
        remove_venv_to_workbook_mapping(workbook_path)
        raise Exception
    
    # handle venv does not exist
    if not Path(found_venv).exists():
        logger.warning(f"the cached venv '{found_venv}' does not exist")
        remove_venv_to_workbook_mapping(workbook_path)
        raise Exception

    return Path(found_venv)



def prompt_user_input(prompt:str) -> str:
    console.print(f"{prompt} ", style=style_prompt_boldface)
    ret = input()
    sys.stdout.flush()
    return ret


def prompt_user_valid_file_path(prompt) -> str:
    v = Path(prompt_user_input(prompt)).resolve()
    while not v.exists():
        console.print(f"file path does not exist: {v}", style=style_error)
        v = Path(prompt_user_input(prompt))
    return str(v)


def prompt_yes_no_input(prompt:str, default:str = "yes") -> str:
    default = default.lower()

    if not default in ["yes", "no"]:
        raise Exception

    lookup = {
        "yes": "yes",
        "y": "yes",
        "no": "no",
        "n": "no",
        "": default.lower()
    }
    def print_prompt():
        console.print(f"{prompt} ", style=style_prompt_boldface, end="")
        console.print(f"[{'Y' if default=='yes' else 'y'}/{'N' if default=='no' else 'n'}]:", style=style_prompt)
        sys.stdout.flush()

    print_prompt()
    inp = input().lower()
    if not (v:=inp.lower()) in lookup.keys():
        if v == "q":
            console.print("User requested to quit. Exiting...", style=style_plain)
            time.sleep(0.5)
            sys.exit()
        console.print(f"{v.lower()} not recognized", style=style_error)
        return prompt_yes_no_input(prompt, default)
    
    ret = lookup[inp]
    if ret == "":
        console.print(default, style=style_plain)
    return lookup[inp]


def get_user_selection(prompt:str, selection_items:list[str], index:int=0) -> str:
    lines_count = "\n".join([prompt] + selection_items).count("\n") + 1

    s_list = []

    def print_selected_line(s:str) -> str:
        console.print(
            (new_s:=(sub_s:="> " + s.replace("\n", "\n  ")) + " " * min(2, get_terminal_width()-len(sub_s))), 
            style=style_selected_option,
        )
        return new_s
    
    def print_generic_line(s:str) -> str:
        console.print(
            (new_s:=(sub_s:="  " + s.replace("\n", "\n  "))), 
            style=style_generic_option,
        )
        return new_s

    def print_options():
        nonlocal s_list
        for i, item in enumerate(selection_items):
            if i == index:
                s_list.append(print_selected_line(item))
            else:
                s_list.append(print_generic_line(item))

    def print_prompt() -> str:
        nonlocal s_list
        console.print(ret:=prompt, style=style_prompt_boldface)
        s_list.append(ret)
    
    print_prompt()
    print_options()
    
    while True:
        clear_text("\n".join(s_list))
        s_list = []
        print_prompt()
        print_options()

        key = readchar.readkey()

        if key == readchar.key.UP:
            index = (index - 1) % len(selection_items)  # Move up
        elif key == readchar.key.DOWN:
            index = (index + 1) % len(selection_items)  # Move down
        elif key == readchar.key.ENTER:
            ret = selection_items[index]
            break
        time.sleep(0.01)

    return ret


def print_info(msg:str):
    console.print("INFO:    ", style=style_plain_boldface, end="")
    console.print(msg, style=style_plain)


def print_success(msg:str):
    console.print("SUCCESS: ", style=style_success_boldface, end="")
    console.print(msg, style=style_success)


def print_warning(msg:str):
    console.print("WARNING: ", style=style_warning_boldface, end="")
    console.print(msg, style=style_warning)


def print_error(msg:str):
    console.print("ERROR:   ", style=style_error_boldface, end="")
    console.print(msg, style=style_error)


def is_py_interpreter_version_match_to_server_config(py_interpreter:Path, workbook_path:Path):
    with open(get_xlpro_workbook_directory(workbook_path) / ".python-version", "r") as f:
        required_version = f.read() 
    active_version = get_py_exe_version(py_interpreter)
    logger.debug(f"required python version for {workbook_path} is {required_version}, active version is {active_version}")
    return active_version == required_version


def dlg_compare_environment_to_requirements_txt(environment_root_path:Path, external_requirements_txt_fp:Path) -> None:
    python_exe = get_python_exe_from_xlpro_root_venv_path(environment_root_path)

    result = subprocess.run(
        [
            "uv",
            "pip",
            "sync",
            "--dry-run",
            "--python",
            str(python_exe.resolve()),
            f"{str(external_requirements_txt_fp.resolve())}",
            f"--find-links={XLPRO_SRC_DIR}",
        ], 
        cwd=str(environment_root_path),
        capture_output=True, 
        # shell=True,
        text=True,
        # check=True,
    )

    lines = result.stderr.split("\n")[:-1]
    if lines[-1] == "Would make no changes":
        print_success(f"requirements.txt match, ok to continue.")
    else:
        # logger.warning(f"requirements.txt does not match.\nmsg:\n{'  '.join(lines)}\n")
        # input("press enter to update the environment per the above")
        print_warning("Virtual environment requirements do not match target, see the following output for installation requirements")
        print(textwrap.indent('\n'.join(lines), '  ') + "\n")
        if (v:=prompt_yes_no_input("Would you like to update the environment per the above changes")) == "yes":
            print_info("Updating the environment dependencies...")
            process = subprocess.Popen(
                [
                    "uv",
                    "pip",
                    "sync",
                    "--python",
                    str(python_exe.resolve()),
                    f"{str(external_requirements_txt_fp.resolve())}",
                    f"--find-links={XLPRO_SRC_DIR}",
                ], 
                cwd=str(environment_root_path),
                text=True,
                stderr=subprocess.PIPE,
                stdout=subprocess.PIPE,
                # check=True,
            )

            def stdout_writer():
                for line in process.stdout:
                    print(line, end='')  # Print each`` line from stdout immediately

            def stderr_writer():
                for line in process.stderr:
                    print(line, end='', file=sys.stderr)  # Print stderr immediately

            import threading
            t1 = threading.Thread(target=stdout_writer)
            t2 = threading.Thread(target=stderr_writer)
            t1.start()
            t2.start()
            process.wait()
            t1.join()
            t2.join()
            print_success("Dependency updates completed successfully.")

            if DEVELOPMENT_INSTALL:
                print_warning("DEVELOPMENT BUILD: Overwriting requirements with editable xlpro version")
                install_editable_default_reqs(py_interpreter_path=python_exe)

        elif v == "no":
            print_warning("Updates skipped due to error, you may be missing requirements for your environment and may need to rectify this manually!")


    return


def dlg_compare_venv_environment_to_required_environment(environment_root_path:Path, workbook_path:Path):
    xlpro_server_dir = get_xlpro_workbook_directory(workbook_path)
    validate_xlpro_folder(xlpro_server_dir)
    required_py_version = read_python_version_file_within_dir(xlpro_server_dir)
    if not environment_root_path.exists():
        print_warning(f"Environment path does not exist '{environment_root_path}'")
        raise FileNotFoundError(f"the environment does not exist, please create another {environment_root_path}")
    active_py_version = get_py_exe_version(get_python_exe_from_xlpro_root_venv_path(environment_root_path))

    print_info(f"Comparing Python version of local environment {environment_root_path} for {workbook_path}...")
    if not compare_py_version_t2(required_py_version, active_py_version):
        print_warning(f"T2 Python versions do not match, active: {active_py_version}, suggested: {required_py_version}")
        # print_error(f"Major warning: T2 Python version mismatch, your environment differs from t")
        # raise Exception("Python version mismatch, please correct. (Risk of overwriting server .python-version is pending development)")
    else:
        print_success(f"T2 Python versions match, ok to continue.")

    dlg_compare_environment_to_requirements_txt(environment_root_path, xlpro_server_dir / "requirements.txt")


def write_requirements_txt_to_folder(xlpro_venv_root_path:Path, xlpro_workbook_dir:Path) -> Path:
    """Writes the requirements txt to an external folder. Provide the parent folder relative to .venv
    for uv to handle dealing with the requirements.txt file.
    Returns the output requirements.txt file because why not"""
    outfile = xlpro_workbook_dir / 'requirements.txt'

    # XXX - todo - check this is writing from the correct interpreter
    result = subprocess.run(
        # todo change this to pip list --format=freeze --not-required!
        [
            "uv",
            "pip",
            "freeze",
            "--python",
            str((xlpro_venv_root_path / ".venv/scripts/python.exe").resolve()),
            ">",
            str((xlpro_workbook_dir / 'requirements.txt').resolve()),
        ],
        # cwd=str(xlpro_venv_root_path.resolve()), # PATH EXISTS
        capture_output=True, 
        shell=True,
        check=True,
        # text=True,
    )
    return outfile


def write_reqs_for_workbook(workbook_path:Path):
    xlpro_venv_root_path = get_valid_venv_root_path_used_for_workbook_from_map(workbook_path)
    if not xlpro_venv_root_path:
        print_error("interpreter does not exist to write requirements.txt, exiting function")
        raise Exception
    xlpro_workbook_dir = get_xlpro_workbook_directory(workbook_path)
    write_requirements_txt_to_folder(
        xlpro_venv_root_path,
        xlpro_workbook_dir,
    )

def write_python_version_file_for_workbook(workbook_path:Path):
    xlpro_venv_root_path = get_valid_venv_root_path_used_for_workbook_from_map(workbook_path)
    if not xlpro_venv_root_path:
        print_error("interpreter does not exist to write requirements.txt, exiting function")
        raise Exception
    xlpro_workbook_dir = get_xlpro_workbook_directory(workbook_path)
    write_python_version_file_for_venv(
        interpreter_path=get_python_exe_from_xlpro_root_venv_path(xlpro_venv_root_path),
        parent_dir=xlpro_workbook_dir,
    )
    pass

import traceback
from functools import wraps

def traceback_log_raise(func):
    @wraps(func)
    def inner(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except SystemExit as e:
            raise
        except KeyboardInterrupt as e:
            raise
        except Exception as e:
            logger.error(f"{func.__qualname__}, error: '{e}'")
            logger.error(f"{traceback.format_exc()}")
            raise
    return inner

def read_json_file_with_comments(fp:Path) -> dict:
    with open(fp, "r", encoding="utf-8") as f:
        txt = f.read()
    # replace the comments with blanks
    txt_no_comments = re.sub(r"\s*(\/\/.*)$", "", txt, flags=re.MULTILINE)
    d = json.loads(txt_no_comments)
    return d


def read_json_file(fp:Path) -> dict:
    with open(fp, "r", encoding="utf-8") as f:
        try:
            d = json.loads(f.read())
        except: 
            d = {}
    return d


def write_json_file(fp:Path, data:dict, indent=2):
    XLPRO_ENVS_DIR.mkdir(exist_ok=True)
    with open(fp, "w", encoding="utf-8") as f:
        f.write(json.dumps(data, indent=indent))
    return
    
    

def read_venv_to_workbooks_map() -> dict[str, list[str]]:
    if not XLPRO_VENV_WORKBOOKS_MAP_JSON_FP.exists():
        write_json_file(XLPRO_VENV_WORKBOOKS_MAP_JSON_FP, {})

    d = read_json_file(XLPRO_VENV_WORKBOOKS_MAP_JSON_FP)
    # for k, v_list in d.items():
    #     if not Path(k).exists():
    #         print_warning(f"xlpro environment {k} does not exist, clearing")
    return d


def write_venv_to_workbooks_mappings_dict(venv_to_workbooks_map:dict) -> None:
    # venv mappings map the environment to the workbook, i.e. environment_path: workbook_path
    write_json_file(XLPRO_VENV_WORKBOOKS_MAP_JSON_FP, venv_to_workbooks_map)


def store_venv_to_workbook_mapping(workbook_path:str|Path, xlpro_venv_parent_path:str|Path):
    """venv to workbook mapping maps the venv_root_dir.parent (xlpro custom name folder) to the workbook filepath
    i.e. one level above the /.venv folder..."""
    
    venv_to_workbooks_map = read_venv_to_workbooks_map()

    # if workbook_path == "default":
    #     str_workbook_path = workbook_path
    #     str_xlpro_venv_parent_path = str_workbook_path
    # else:
    str_workbook_path = str(workbook_path.resolve())
    str_xlpro_venv_parent_path = str(xlpro_venv_parent_path.resolve())


    workbook_path = None
    xlpro_venv_parent_path = None

    # construct reversed dict to check if the workbook isn't linked to a venv
    venv_to_workbooks_map_reversed:dict[str, str] = {}
    for k, v_list in venv_to_workbooks_map.items():
        for v in v_list:
            if v in venv_to_workbooks_map_reversed.keys():
                raise Exception("duplicated value, should not be possible")
            venv_to_workbooks_map_reversed[str(v)] = str(k)

    # remove the existing venv link if it exists per the above dict
    if str_workbook_path in venv_to_workbooks_map_reversed.keys():
        logger.warning(f"workbook {str_workbook_path} was previously mapped to {venv_to_workbooks_map_reversed[str_workbook_path]}")
        venv_to_workbooks_map[venv_to_workbooks_map_reversed[str_workbook_path]].remove(str_workbook_path)

    # add the new venv to workbook link as a new list item
    if not str_xlpro_venv_parent_path in venv_to_workbooks_map.keys():
        venv_to_workbooks_map[str_xlpro_venv_parent_path] = []
    venv_to_workbooks_map[str_xlpro_venv_parent_path].append(str_workbook_path)

    # save out the dict
    write_venv_to_workbooks_mappings_dict(venv_to_workbooks_map)


def remove_venv_to_workbook_mapping(workbook_path:str|Path, venv_path:str|Path|None=None) -> None:
    venv_to_workbooks_map = read_venv_to_workbooks_map()
    workbook_path = Path() / workbook_path
    str_workbook_path = str(workbook_path.resolve())
    str_venv_path = str(venv_path.resolve())

    # if only the workbook is given, delete the workbook and its links. E.g. the user wants to unlink a workbook
    if workbook_path and (venv_path is None):
        del venv_to_workbooks_map[str_workbook_path]
        write_venv_to_workbooks_mappings_dict()
        return
    # if only the venv is provided, remove any occurrences of the venv. E.g. it was deleted manually
    elif venv_path and (workbook_path is None):
        # todo - XXX - test this code
        for k, v, in venv_to_workbooks_map.items():
            if str_venv_path in v:
                venv_to_workbooks_map[k].remove(str_venv_path)
        write_venv_to_workbooks_mappings_dict()
        return
    
    workbook_path = None
    venv_path = None

    # handle specific venv and workbook link removal
    if not str_venv_path in venv_to_workbooks_map.keys():
        logger.error(f"venv_path {str_venv_path} not found in {XLPRO_VENV_WORKBOOKS_MAP_JSON_FP.name}")
        return
    if not str_workbook_path in venv_to_workbooks_map[str_venv_path]:
        logger.error(f"workbook_path {str_workbook_path} not found in {XLPRO_VENV_WORKBOOKS_MAP_JSON_FP.name} key {str_venv_path}")

    venv_to_workbooks_map[str_venv_path].remove(str_workbook_path)
    write_venv_to_workbooks_mappings_dict(venv_to_workbooks_map)


def get_valid_venv_root_path_used_for_workbook_from_map(workbook_path:str|Path) -> None|Path:
    """Finds the venv last used for the specified workbook if it exists, else returns None"""
    str_workbook_path = str(workbook_path.resolve())
    workbook_path = None
    venv_to_workbooks_map = read_venv_to_workbooks_map()
    # venv_to_workbooks_map_reversed is {workbook_path: environment_path}
    venv_to_workbooks_map_reversed = {v: k for k, v_list in venv_to_workbooks_map.items() for v in v_list}
    if not str_workbook_path in venv_to_workbooks_map_reversed.keys():
        logger.error(f"workbook_path {str_workbook_path} was not found within in {XLPRO_VENV_WORKBOOKS_MAP_JSON_FP.name}. No environment found")
        return None

    venv_root_path = Path() / venv_to_workbooks_map_reversed[str_workbook_path]

    # do not allow a non-existent venv out of this function!
    if not venv_root_path.exists():
        print_warning(f"path does not exist '{venv_root_path}', removing the link for '{workbook_path}'")
        remove_venv_to_workbook_mapping(str_workbook_path, venv_path=venv_root_path)
        return

    return venv_root_path

def remove_venv_mapping_for_workbook(workbook_path:Path, dialogue=False):
    venv_path = get_valid_venv_root_path_used_for_workbook_from_map(workbook_path=workbook_path)
    if venv_path is None:
        print_info(f"No venv is currently mapped to {workbook_path.name}. No further actions required.")
        return
    if dialogue:
        if prompt_yes_no_input(f"Are you sure you want to remove venv link for '{workbook_path.name}'", default="yes") == "no":
            print_info("User requested to not proceeed.")
            return
    print_info(f"Removing venv-workbook link '{venv_path.name}'-'{workbook_path.name}'")
    remove_venv_to_workbook_mapping(workbook_path=workbook_path, venv_path=venv_path)


def delete_stale_environments():
    """Deletes all venvs not currently tied to workbooks"""
    raise NotImplementedError


def write_settings_json_python_path(parent_dir:Path, absolute_python_exe_path:Path) -> Path:
    if not absolute_python_exe_path.is_absolute():
        raise Exception(f"python exe path must be absolute (in order to preserve symlinks)")
    vscode_dir = (parent_dir / ".vscode")
    (parent_dir / ".vscode").mkdir(exist_ok=True)
    settings_dict = {}
    settings_dict["python.defaultInterpreterPath"] = str(absolute_python_exe_path)
    settings_json_path = vscode_dir / "settings.json"

    py_def_path_str = textwrap.indent(
        textwrap.dedent(
            ("""
                // XLPRO DEFAULT INTERPRETER PATH - START
                "python.defaultInterpreterPath": "{absolute_python_exe_path}"
                // XLPRO DEFAULT INTERPRETER PATH - END
                """
            )[1:].format(absolute_python_exe_path=str(absolute_python_exe_path).replace("\\", "\\\\"))
        ),
        prefix=" " * 4
    )

    default_settings_json_text = (
        textwrap.dedent(
            """
            {{
                // This file was auto-generated by xlpro-cli.exe
            {py_def_path_str}
            }}"""
        )
    )[1:].format(py_def_path_str=py_def_path_str)

    def _write_settings_json_content(content:str):
        with open(settings_json_path, "w") as f:
            f.write(content)

    if settings_json_path.exists():
        print_warning("settings.json already exists, substituting xlpro default interpreter path block...")
        with open(settings_json_path, "r") as f:
            settings_json_txt = f.read()
        mtch = re.search(r"\n\s*\/\/ XLPRO DEFAULT INTERPRETER PATH - START.*\/\/ XLPRO DEFAULT INTERPRETER PATH - END", settings_json_txt, re.DOTALL)

        # if a quotation mark is found, it needs a trailing comma
        mtch2 = re.search(r"\n\s*\/\/ XLPRO DEFAULT INTERPRETER PATH - START.*\/\/ XLPRO DEFAULT INTERPRETER PATH - END.*?\"", settings_json_txt, re.DOTALL)

        if not mtch:
            print_error("settings.json incorrectly configured for substitution")
            if prompt_yes_no_input("Force overwrite launch.json with defaults?", default="no") == "yes":
                print_info("Writing default settings.json")
                _write_settings_json_content(default_settings_json_text)
                print_success(f"{settings_json_path} written")
                return

        injected_json_content = settings_json_txt.replace(mtch.group(), "\n" + py_def_path_str)

        # if a quotation mark is found, it needs a trailing comma
        mtch2 = re.search(r"\n\s*\/\/ XLPRO DEFAULT INTERPRETER PATH - START.*\/\/ XLPRO DEFAULT INTERPRETER PATH - END.*?\"", settings_json_txt, re.DOTALL)
        trailing_comma = bool(mtch2)
        if trailing_comma:
            injected_json_content = re.sub(r'(XLPRO DEFAULT INTERPRETER PATH - START.*?"python.defaultInterpreterPath".*?)\n', r"\1,\n", injected_json_content, flags=re.DOTALL)

        print_info("Writing injected json content to settings.json")
        _write_settings_json_content(injected_json_content)
    else:
        print_info("Writing default settings.json")
        _write_settings_json_content(default_settings_json_text)

    print_success(f"{settings_json_path} written")

    # write_json_file(settings_json_path, settings_dict)
    logger.warning(f"created settings.json at {settings_json_path}. Ensure you run the 'Python: Clear Workspace Interpreter' command from the command pallette")
    return settings_json_path


def write_launch_json(parent_dir:Path, debugpy_port:int=5678):
    if not isinstance(debugpy_port, int):
        raise TypeError

    xlpro_configuration_string = textwrap.indent(
        textwrap.dedent("""
            // XLPRO DEBUG CONFIGURATION - START
            "name": "xlpro debugpy",
            "type": "debugpy",
            "request": "attach",
            "subProcess": true,
            "connect": {{
                "host": "localhost",
                "port": {debugpy_port}
            }},
            "pathMappings": [
                {{
                    "localRoot": "${{workspaceFolder}}",
                    "remoteRoot": "${{workspaceFolder}}"
                }}
            ],
            "justMyCode": true
            // XLPRO DEBUG CONFIGURATION - END"""[1:].format(debugpy_port=debugpy_port)
        ),
        prefix=" " * 12
    )

    default_launch_json_content = (
        textwrap.dedent(
            """
            {{
                // This file was auto-generated by xlpro-cli.exe
                "version": "0.2.0",
                "configurations": [
                    {{
            {xlpro_configuration_string}
                    }}
                ]
            }}"""
        ).format(
            xlpro_configuration_string=xlpro_configuration_string,
        )
    )
    pass

    launch_json_path = parent_dir / ".vscode/launch.json"

    def _write_launch_json_content(content):
        with open(launch_json_path, "w") as f:
            f.write(content)
            
    if launch_json_path.exists():
        print_warning("launch.json already exists, substituting xlpro debug configuration block...")
        with open(launch_json_path, "r") as f:
            launch_json_txt = f.read()
        mtch = re.search(r"\n\s*\/\/ XLPRO DEBUG CONFIGURATION - START.*\/\/ XLPRO DEBUG CONFIGURATION - END", launch_json_txt, re.DOTALL)
        if not mtch:
            print_error("launch.json incorrectly configured for substitution")
            if prompt_yes_no_input("Force overwrite launch.json with defaults?", default="no") == "yes":
                print_info("Writing default launch.json")
                _write_launch_json_content(default_launch_json_content)
                print_success(f"{launch_json_path} written")
                return

        injected_json_content = launch_json_txt.replace(mtch.group(), "\n" + xlpro_configuration_string)
        print_info("Writing injected json content to launch.json")
        _write_launch_json_content(injected_json_content)
    else:
        print_info("Writing default launch.json")
        _write_launch_json_content(default_launch_json_content)

    print_success(f"{launch_json_path} written")


def write_python_version_file_for_venv(interpreter_path:Path, parent_dir:Path):
    with open(parent_dir / ".python-version", "w") as f:
        f.write(get_py_exe_version(interpreter_path))


def python_version_string_is_correctly_formatted(s:str) -> bool:
    version_match = re.match(r"(\d+\.\d+\.\d+)", s)
    return bool(version_match)


def read_python_version_file_within_dir(parent_dir:Path) -> str:
    if not(fp:=(parent_dir / ".python-version")).exists():
        raise FileNotFoundError(f"python version file not found at {fp}")
    with open(fp, "r") as f:
        ret = f.read()
    if not python_version_string_is_correctly_formatted(ret):
        raise Exception(f"Python version string {ret} is not correctly formatted within {fp}")
    return ret


def get_venv_root_directory_for_xlpro(venv_root_path:Path) -> Path:
    # XXX - todo - replace this with the .venv file
    if (venv_root_path / ".venv/scripts/python.exe").exists():
        return venv_root_path
    return get_venv_root_directory(venv_root_path).parent
    if (venv_root_path / ".venv/scripts/python.exe").exists():
        ret = venv_root_path
    elif (venv_root_path / "scripts/python.exe").exists():
        ret = venv_root_path.parent
    elif venv_root_path.name == "python.exe":
        ret = venv_root_path.parent.parent.parent
    else:
        raise Exception("invalid venv root path provided")
    return ret


def get_venv_root_directory(venv_root_path:Path) -> Path:
    """returns the <dirname> of a virtual environment for a semi-arbitrary input patg
        <dirname>/scripts/python.exe
    """
    if (venv_root_path / "scripts/python.exe").exists():
        ret = venv_root_path
    elif (venv_root_path / "python.exe").exists():
        ret = venv_root_path.parent.parent
    elif venv_root_path.name == "python.exe":
        ret = venv_root_path.parent.parent
    else:
        raise Exception("invalid venv root path provided")
    return ret


def write_server_vscode_config_settings(workbook_path:Path, active_venv:Path):
    """Writes the following data to path/to/book.xlsx/../book.xlsx/
        - writes ./.vscode/settings.json
        - writes ./.vscode/launch.json
    Server environment files are read to configure the environments used to execute python code for this workbook.
    """

    active_venv_standardized_fp = get_venv_root_directory_for_xlpro(active_venv)
    xlpro_server_dir = get_xlpro_workbook_directory(workbook_path=workbook_path)
    initialize_and_get_workspace_xlpro_dir(workbook_path=workbook_path)
    # write settings.json to server location for IDE integration
    venv_exe_path = get_python_exe_from_xlpro_root_venv_path(active_venv_standardized_fp)
    
    write_settings_json_python_path(xlpro_server_dir, venv_exe_path)
    # write launch.json for debug server support
    write_launch_json(xlpro_server_dir)


def write_server_environment_settings(workbook_path:Path, active_venv:Path):
    """Writes the following data to path/to/book.xlsx/../book.xlsx/
        - writes requirements.txt
        - writes .python-version
    Server environment files are read to configure the environments used to execute python code for this workbook.
    """

    active_venv_standardized_fp = get_venv_root_directory_for_xlpro(active_venv)
    xlpro_server_dir = get_xlpro_workbook_directory(workbook_path=workbook_path)

    # initialize_and_get_workspace_xlpro_dir(workbook_path=workbook_path)

    # write requirements.txt
    write_requirements_txt_to_folder(active_venv_standardized_fp, xlpro_server_dir)
    # write .python-version
    write_python_version_file_for_venv(get_python_exe_from_xlpro_root_venv_path(active_venv_standardized_fp), xlpro_server_dir)



def write_local_environment_settings(active_venv:Path):
    """Writes the following data to path/to/local/xlpro/venv/XXX/.venv/../.xlpro/
        - requirements.txt
        - .python-version
    """
    raise NotImplementedError("There should be no need to write this data to this directory except for debugging purposes I suppose.")
    
    active_venv_standardized_fp = get_venv_root_directory_for_xlpro(active_venv)
    active_venv_xlpro_dir = active_venv_standardized_fp / ".xlpro"
    # write requirements.txt
    write_requirements_txt_to_folder(active_venv_standardized_fp, active_venv_xlpro_dir)
    # write .python-version
    write_python_version_file_for_venv(get_python_exe_from_xlpro_root_venv_path(active_venv_standardized_fp), active_venv_xlpro_dir)


def write_local_venv_workbook_link_data(workbook_path:Path, active_venv:Path):
    """updates the venv cache, removes the mapping between the existing venv and the workbook if it exists, otherwise
    it just adds the new environment to the dict."""
    active_venv_standardized_fp = get_venv_root_directory_for_xlpro(active_venv)

    existing_venv = get_valid_venv_root_path_used_for_workbook_from_map(workbook_path=workbook_path)
    if existing_venv is not None:
        print_info(f"overwriting existing link for {workbook_path} to {existing_venv}")
        remove_venv_to_workbook_mapping(workbook_path=workbook_path, venv_path=existing_venv)
        store_venv_to_workbook_mapping(workbook_path=workbook_path, xlpro_venv_parent_path=active_venv_standardized_fp)
        return
    
    print_info(f"no existing interpreter found for {workbook_path}")
    print_info(f"creating new link between {workbook_path} and {active_venv_standardized_fp}")
    store_venv_to_workbook_mapping(workbook_path=workbook_path, xlpro_venv_parent_path=active_venv_standardized_fp)


class venv_types:
    SYSTEM_INTERPRETER = "system"
    REUSED_LOCAL_VENV = "reused"
    UV_DOWNLOAD_NEW_VENV = "uv-download"
    REUSED_XLPRO_VENV = "reused-xlpro"



def get_recommended_python_version() -> str:
    with open(fp:=(XLPRO_ROOT_PATH / ".python-version-recommended")) as f:
        data = f.read()

    if not re.match(r"3\.\d+(?:\.\d+)?", data, re.IGNORECASE):
        raise ValueError(f"Provided recommended '{data}' version is invalid in {fp}")
    
    return data


class MenuSelection:
    BASIC = 1
    ADVANCED = 2


def dlg_select_and_optionally_create_valid_python_interpreter(version_required=None, allow_override:bool=False, menutype=MenuSelection.BASIC) -> tuple[Path, venv_types]:

    # result = subprocess.run("where.exe python", shell=True, capture_output=True, text=True, check=True)
    result = subprocess.run(["where.exe", "python"], capture_output=True, text=True, check=True)
    py_path_locations_system = result.stdout.split("\n")[:-1]

    menu_items_advanced = []
    menu_items_advanced += [
        (default_prefix:="(recommended)") +  f" Create New Python {get_recommended_python_version()} Environment" 
    ]
    menu_items_basic = menu_items_advanced.copy()
    menu_items_basic += [
        (enable_advanced:="(advanced)    Enable advanced selection")
    ]
    menu_items_advanced += [
        f"(system)      {x}" for x in py_path_locations_system
    ]
    xlpro_exes = get_xlpro_python_interpreters()
    local_version_str1 = "(xlpro-local)"
    menu_items_advanced += [
        f"{local_version_str1} {x}" for x in xlpro_exes
    ]
    menu_items_basic += [
        f"{local_version_str1} {x}" for x in xlpro_exes
    ]
    uv_py_exes = get_uv_python_interpreters() 
    menu_items_advanced += [
        f"(uv-local)    {x}" for x in uv_py_exes
    ]
    uv_dl_prefix='(uv-download)'
    menu_items_advanced += [
        # f"{uv_dl_prefix} Python 3.12",
        # f"{uv_dl_prefix} Python 3.13",
        # f"{uv_dl_prefix} Python 3.14",
        other_version_str:="(uv-download) Specify Version",
        local_version_str2:="(local)       Reuse other [specify path]",
    ]

    
    override_str = "(override)    Allow overriding of the Python version [ADVANCED USE ONLY]"
    nooverride_str = "(override)    Remove overriding"
    if not allow_override:
        menu_items_advanced.append(override_str)
        menu_items_basic.append(override_str)
    else:
        menu_items_advanced.append(nooverride_str)
        menu_items_basic.append(nooverride_str)

    disable_advanced = "(simple)      Go back to simple selection"
    menu_items_advanced.append(disable_advanced)

    d = {
        MenuSelection.BASIC: menu_items_basic,
        MenuSelection.ADVANCED: menu_items_advanced,
    }

    menu_items = d[menutype]
    

    def _convert_menu_item_to_path(s:str):
        return re.match(r"^\(.+\)\s+(.*)$", s).group(1)

    prompt = "Select your python interpreter" + ("" if version_required is None else f"[requires {version_required}]") + (" [OVERRIDES ALLOWED]" if allow_override else "")

    ret = get_user_selection(prompt=prompt, selection_items=menu_items)



    # Handle the override toggle
    if ret == override_str:
        print_warning("User has requested to override the python version safeguards, mismatched versions can be forced on the next prompt")
        return dlg_select_and_optionally_create_valid_python_interpreter(version_required=version_required, allow_override=True)
    # Handle the override toggle
    if ret == nooverride_str:
        return dlg_select_and_optionally_create_valid_python_interpreter(version_required=version_required, allow_override=False)

    # Enable advanced configuration
    if ret == enable_advanced:
        print_warning("Advanced configuration selected")
        return dlg_select_and_optionally_create_valid_python_interpreter(version_required=version_required, allow_override=allow_override, menutype=MenuSelection.ADVANCED)
    
    # Enable advanced configuration
    if ret == disable_advanced:
        print_warning("Advanced configuration selected")
        return dlg_select_and_optionally_create_valid_python_interpreter(version_required=version_required, allow_override=allow_override, menutype=MenuSelection.BASIC)


    do_proceed_with_creation = False
    rettype:venv_types = None

    # handle uv other version
    if ret == other_version_str:
        uv_py_version = prompt_user_input(msg:="Provide a python version to download e.g. 3.12.2, (b to go back)")
        if uv_py_version.lower() == "b":
            return dlg_select_and_optionally_create_valid_python_interpreter(version_required=version_required, allow_override=allow_override, menutype=menutype)
        if version_required is not None:
            if not compare_py_version_t2(uv_py_version, version_required):
                print_error(f"provided version {uv_py_version} is not compatible with {version_required}")
                if allow_override:
                    if prompt_yes_no_input("OVERRIDE ENABLED: Do you want to override the python version? [ADVANCED USE ONLY]") == "yes":
                        do_proceed_with_creation = True
            else:
                do_proceed_with_creation = True

            if not do_proceed_with_creation:
                return dlg_select_and_optionally_create_valid_python_interpreter(version_required=version_required, allow_override=allow_override, menutype=menutype)
                
            
        ret = uv_download_python_version(uv_py_version)
        rettype = venv_types.UV_DOWNLOAD_NEW_VENV


    # handle prescribed uv-download request specific download request 
    # (must occur after handling uv other version for namespace clash)
    elif ret.startswith(uv_dl_prefix):
        provided_py_version = re.search(r"Python (\d\.\d+)", ret).group(1)
        if version_required is not None:
            if not compare_py_version_t2(provided_py_version, version_required):
                print_error(f"provided version {provided_py_version} is not compatible with {version_required}")
                if allow_override:
                    if prompt_yes_no_input("OVERRIDE ENABLED: Do you want to override the python version? [ADVANCED USE ONLY]") == "yes":
                        do_proceed_with_creation = True
            else:
                do_proceed_with_creation = True
            if not do_proceed_with_creation:
                return dlg_select_and_optionally_create_valid_python_interpreter(version_required=version_required, allow_override=allow_override, menutype=menutype)
                    
        ret = uv_download_python_version(provided_py_version)
        rettype = venv_types.UV_DOWNLOAD_NEW_VENV        



    # Download the recommended python version
    elif ret.startswith(default_prefix):
        provided_py_version = re.search(r"Python (\d\.\d+)", ret).group(1)
        # m = read_venv_to_workbooks_map()
        # val0 = m.get("default", None)
        # if val0 is not None:
        #     if (val1:=Path(val0)).exists():
        #         print_info("Default interpreter exists, resorting to this one.")
        #         ret = val1
        #         rettype = venv_types.REUSED_XLPRO_VENV
        #         return (ret, rettype)

        if version_required is not None:
            if not compare_py_version_t2(provided_py_version, version_required):
                print_error(f"provided version {provided_py_version} is not compatible with {version_required}")
                if allow_override:
                    if prompt_yes_no_input("OVERRIDE ENABLED: Do you want to override the python version? [ADVANCED USE ONLY]") == "yes":
                        do_proceed_with_creation = True
            else:
                do_proceed_with_creation = True

            if not do_proceed_with_creation:
                return dlg_select_and_optionally_create_valid_python_interpreter(version_required=version_required, allow_override=allow_override, menutype=menutype)
        
        
        ret = uv_download_python_version(provided_py_version)
        # store_venv_to_workbook_mapping("default", ret)
        rettype = venv_types.UV_DOWNLOAD_NEW_VENV




    # XXX - todo - handle reuse of an xlpro venv...
    # if the user selects an existing xlpro venv, we don't want to create a new virtual environment.
    elif ret.startswith(local_version_str1):
        ret = _convert_menu_item_to_path(ret)
        ret = Path(ret)      

        if not ret.exists():
            raise FileNotFoundError(f"The file does not exist at '{ret}'. Aborting")
        
        provided_py_version = get_py_exe_version(ret)
        if version_required is not None:
            if not compare_py_version_t2(provided_py_version, version_required):
                print_error(f"Provided version {provided_py_version} is not compatible with {version_required}")
                if allow_override:
                    if prompt_yes_no_input("OVERRIDE ENABLED: Do you want to override the python version? [ADVANCED USE ONLY]") == "yes":
                        do_proceed_with_creation = True

            if not do_proceed_with_creation:
                return dlg_select_and_optionally_create_valid_python_interpreter(version_required=version_required, allow_override=allow_override, menutype=menutype)
        
        # ret = py_path
        rettype = venv_types.REUSED_XLPRO_VENV

    # handle mapping to a local virtual environment.
    elif ret == local_version_str2:
        py_path = Path() / prompt_user_input(msg:="Provide a path to a local virtual environment python.exe, (b to go back)")
        check, err = is_interpreter_valid_venv_and_exists(py_path)
        while not check:
            print_error(f"Interpreter path is invalid: {str(err)}")
            if prompt_user_input(msg).lower() == "b":
                return dlg_select_and_optionally_create_valid_python_interpreter(version_required=version_required, allow_override=allow_override, menutype=menutype)
            py_path = Path() / prompt_user_input(msg)
            check, err = is_interpreter_valid_venv_and_exists(py_path)


        provided_py_version = get_py_exe_version(py_path)
        if version_required is not None:
            if not compare_py_version_t2(provided_py_version, version_required):
                print_error(f"Provided version {provided_py_version} is not compatible with {version_required}")
                if allow_override:
                    if prompt_yes_no_input("OVERRIDE ENABLED: Do you want to override the python version? [ADVANCED USE ONLY]") == "yes":
                        do_proceed_with_creation = True

            if not do_proceed_with_creation:
                return dlg_select_and_optionally_create_valid_python_interpreter(version_required=version_required, allow_override=allow_override, menutype=menutype)

        ret = py_path
        rettype = venv_types.REUSED_LOCAL_VENV

    # any of the automatically fetched environments? at least uv-local environments.
    else:
        ret = _convert_menu_item_to_path(ret)
        ret = Path(ret)        
        if not ret.exists():
            raise FileNotFoundError(f"the file does not exist {ret}")
        
        provided_py_version = get_py_exe_version(ret)
        if version_required is not None:
            if not compare_py_version_t2(provided_py_version, version_required):
                print_error(f"Provided version {provided_py_version} is not compatible with {version_required}")
                if allow_override:
                    if prompt_yes_no_input("OVERRIDE ENABLED: Do you want to override the python version? [ADVANCED USE ONLY]") == "yes":
                        do_proceed_with_creation = True

            if not do_proceed_with_creation:
                return dlg_select_and_optionally_create_valid_python_interpreter(version_required=version_required, allow_override=allow_override, menutype=menutype)
        rettype = venv_types.SYSTEM_INTERPRETER

    if rettype is None:
        raise Exception
    
    if not isinstance(ret, Path):
        raise TypeError

    return (ret, rettype)
    

def install_production_default_reqs_get_process(py_interpreter_path):
    reqs = [
        "pip",
        "xlpro",
        # f"{str(path_to_xlpro_whl)}",
        # f"xlpro=={xlpro.__version__}",
    ]
    process = subprocess.Popen(
        [
            # "xlpro-server.exe"
            "uv",
            "pip",
            "install",
            f"--find-links={XLPRO_SRC_DIR}",
            "--python",
            str(py_interpreter_path),
        ] + reqs,
        # check=True,
        # capture_output=True,
        text=True,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    
    return process

def install_editable_default_reqs(py_interpreter_path):
    """Install editable """
    process = subprocess.Popen(
        [
            # "xlpro-server.exe"
            "uv",
            "pip",
            "install",
            "--python",
            str(py_interpreter_path),
            "pip"
        ],
        # check=True,
        # capture_output=True,
        text=True,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    for line in process.stdout:
        print(line, end='')  # Print each line from stdout immediately
    for line in process.stderr:
        print(line, end='', file=sys.stderr)  # Print stderr immediately
    # Wait for the subprocess to finish
    process.wait()

    # path_to_xlpro = Path(r"C:\Users\Daniel Evans\projects\xlpro\xlpro_module")
    path_to_xlpro = Path(__file__).parent.parent.parent / "xlpro_module"
    process = subprocess.Popen(
        [
            # "xlpro-server.exe",
            str(py_interpreter_path),
            "-m"
            "pip",
            "install",
            "-e",
            f"{str(path_to_xlpro)}",
        ],
        # check=True,
        # capture_output=True,
        text=True,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    for line in process.stdout:
        print(line, end='')  # Print each line from stdout immediately
    for line in process.stderr:
        print(line, end='', file=sys.stderr)  # Print stderr immediately
    # Wait for the subprocess to finish
    process.wait()


def install_requirements(py_interpreter_path:Path, requirements:list[str]):
    if not py_interpreter_path.is_absolute():
        raise Exception("path must be absolute")
    
    process = None

    def stdout_writer():
        for line in process.stdout:
            print(line, end='')  # Print each`` line from stdout immediately

    def stderr_writer():
        for line in process.stderr:
            print(line, end='', file=sys.stderr)  # Print stderr immediately


    if DEVELOPMENT_INSTALL:
        print_warning("DEVELOPMENT RELEASE development default requirements")
        install_editable_default_reqs(py_interpreter_path=py_interpreter_path)
        print_success("DEVELOPMENT RELEASE default requirements install complete.")
    else:
        import threading    
        print_info("Installing default requirements")
        process = install_production_default_reqs_get_process(py_interpreter_path=py_interpreter_path)
        t1 = threading.Thread(target=stdout_writer)
        t2 = threading.Thread(target=stderr_writer)

        t1.start()
        t2.start()

        # Wait for the subprocess to finish
        process.wait()

        t1.join()
        t2.join()
        print_success("Default requirements install complete.")

    pass

def get_xlpro_whl_fp():
    fps = list(XLPRO_SRC_DIR.glob("*.whl"))
    if len(fps)> 1:
        raise Exception("multiple wheels found, abandoning")
    if not fps:
        raise Exception("no wheels found, abandoning")
    return fps[0]

def install_default_requirements(py_interpreter_path:Path):
    # path_to_xlpro = Path().resolve()
    # path_to_xlpro = Path(r"C:\Users\Daniel Evans\projects\xlpro\xlpro_module")
    # path_to_xlpro = get_xlpro_whl_fp()
    # reqs = [
    #     "pip",
    #     f"-e \"{str(path_to_xlpro)}\"",
    # ]
    reqs = []
    print_info(f"Installing default requirements for {py_interpreter_path}")
    install_requirements(py_interpreter_path=py_interpreter_path, requirements=reqs)
    print_success(f"Default requirements installed for {py_interpreter_path}")
    

def dlg_xlpro_initialize_workbook(workbook_path:Path):
    """dialogue run when initializing a workbook. user is prompted to create a new virtual environment if the current one is not compatible
    nb: compatibility checks are crude, only checks if the x.xx python version string is a match."""

    def dlg_user_selects_or_creates_valid_interpreter(version_required=None) -> Path:
        interpreter_path, request_type = dlg_select_and_optionally_create_valid_python_interpreter(version_required)
        print_info(f"user requested venv type: {request_type}")

        # handle a virtual environment created from an existing system interpreter or a newly downloaded uv interpreter.
        if request_type in [venv_types.SYSTEM_INTERPRETER, venv_types.UV_DOWNLOAD_NEW_VENV]:
            print_info(f"creating new virtual environment from {interpreter_path}...")
            xlpro_venv_root_path = create_xlpro_venv_from_interpreter_and_get_root_path(interpreter_path)
            print_success(f"virtual environment creation successful at {xlpro_venv_root_path}")
        
        # handle a reused symlinked venv (creates a dummy environment pointing to the original system venv directory)
        elif request_type == venv_types.REUSED_LOCAL_VENV:
            print_info(f"registering symlinked interpreter to {interpreter_path}...")
            user_specified_venv_root_dir = get_venv_root_directory(interpreter_path)
            xlpro_venv_path = create_symbolic_venv_from_existing_venv(existing_venv_root_path=user_specified_venv_root_dir)
            xlpro_venv_root_path = get_venv_root_directory_for_xlpro(xlpro_venv_path)
            print_success(f"symlinked venv maps {interpreter_path} to {xlpro_venv_root_path} creation successful")
        
        # handle a reused xlpro venv (does not create a new environment)
        elif request_type == venv_types.REUSED_XLPRO_VENV:
            print_info(f"re-using existing xlpro interpreter {interpreter_path}...")
            user_specified_venv_root_dir = get_venv_root_directory(interpreter_path)
            xlpro_venv_root_path = get_venv_root_directory_for_xlpro(user_specified_venv_root_dir)

        return xlpro_venv_root_path
    
    # def install_requirements():
    #     # XXX - todo - this shouldn't be a catch all, add logic at some point...
    #     xlpro_py_exe = get_python_exe_from_xlpro_root_venv_path(xlpro_venv_root_path)
    #     install_default_requirements(xlpro_py_exe)

    #     # update the local venv mappings
    #     print_info(f"Updating local venv for {workbook_path} and {xlpro_venv_root_path}...")
    #     write_local_venv_workbook_link_data(workbook_path=workbook_path, active_venv=xlpro_venv_root_path)
    #     print_success(f"Updating local venv completed successfully.")

    #     # update the workbook .xlpro directory metadata
    #     xlpro_on_save_to_server(workbook_path=workbook_path)

    #     # return the environment path
    #     new_venv_path = xlpro_venv_root_path
    #     return new_venv_path
    

    is_xlpro = is_existing_xlpro_workbook_folder(workbook_path=workbook_path)
    xlpro_server_dir = get_xlpro_workbook_directory(workbook_path=workbook_path)
    write_env_settings = False
    # if it is already an xlpro file, 
    # prompt the user to initialize their own environment if one does not already exist
    # the user has several options to create a new virtual environment from an intepreter, or map to an existing virtual environment.
    if is_xlpro:
        validate_xlpro_folder(xlpro_server_dir)
        print_warning("Workbook appears to already be configured for xlpro")
        venv_root_path = get_valid_venv_root_path_used_for_workbook_from_map(workbook_path)
        xlpro_recommended_py_version = read_python_version_file_within_dir(xlpro_server_dir)
        print_info(f"the recommended py version for this workbook is {xlpro_recommended_py_version}")
        # if the virtual environment does not exist on the user's machine, we must create a new one.
        if venv_root_path is None:
            print_warning(f"Could not locate environment a user environment registered with this workbook")
            if prompt_yes_no_input("Would you like to initialise a new python environment for this workbook?") == "yes":
                # We should do the following
                # 1. install the correct python version (let the user choose)
                # 2. this is an xlpro file, read the requirements, try and compare the environments and fall back to the default requirements.
                g = re.search(r'(\d\.\d+)', xlpro_recommended_py_version).group(1)
                print_warning(f"Please select a Python version matching: '{g}.*'")
                venv_root_path = dlg_user_selects_or_creates_valid_interpreter(version_required=get_t2_version_str(xlpro_recommended_py_version))
            else:
                print_warning("No further actions to take. Exiting...")
                sleep_then_close()

        # walk the user through comparing the venv environment to the required.
        # do not push anything to the server requirements! Let the user do this manually.
        print_info("Comparing venv requirements to required...")
        dlg_compare_venv_environment_to_required_environment(environment_root_path=venv_root_path, workbook_path=workbook_path)
        print_info("Comparison complete.")
        print_info("No modifications have been made to the server environment settings. The author should push these via the add-in.")

    # if its not an xlpro directory, we can start a new venv for it
    # the user has the same options to create the environment as above.
    else:
        print_info(f"Workbook is not initialized for xlpro, the following steps will configure your environment, no version required")
        venv_root_path = dlg_user_selects_or_creates_valid_interpreter(version_required=None)
        #i install the default requirements
        py_interpreter_path = get_python_exe_from_xlpro_root_venv_path(venv_root_path)
        install_default_requirements(py_interpreter_path=py_interpreter_path)
        write_env_settings = True


    print_info(f"Updating local venv for {workbook_path} and {venv_root_path}...")
    write_local_venv_workbook_link_data(workbook_path=workbook_path, active_venv=venv_root_path)
    write_server_vscode_config_settings(workbook_path=workbook_path, active_venv=venv_root_path)
    if write_env_settings:
        write_server_environment_settings(workbook_path=workbook_path, active_venv=venv_root_path)
    print_success(f"Updating local venv completed successfully.")


def validate_xlpro_folder(fp:Path):
    errors = []

    # .python-version exists and is correctly formatted.
    try:
        read_python_version_file_within_dir(fp)
    except Exception as e:
        errors.append(e)

    # requirements.txt
    if not (v:=(fp / "requirements.txt")).exists():
        e = FileNotFoundError(f"Requirements does not exist. {v}")
        errors.append(e)

    # functions.py
    if not (v:=(fp / "functions.py")).exists():
        e = FileNotFoundError(f"functions.py does not exist. {v}")
        errors.append(e)
    
    # subroutines.py
    if not (v:=(fp / "subroutines.py")).exists():
        e = FileNotFoundError(f"subroutines.py does not exist. {v}")
        errors.append(e)

    for e in errors:
        logger.error(e)

    if errors:
        raise Exception("xlpro folder is invalid, please rectify per the error messages")
    

    


    


def xlpro_on_save_to_server(workbook_path:Path):
    """code run to save the environment configuration to the server location"""
    xlpro_venv_root_path = get_valid_venv_root_path_used_for_workbook_from_map(workbook_path)
    print_info(f"Writing server files for {workbook_path} and {xlpro_venv_root_path}...")
    write_server_environment_settings(workbook_path=workbook_path, active_venv=xlpro_venv_root_path)
    print_success(f"Writing server files completed successfully.")


def xlpro_change_workbook_venv(workbook_path:Path):
    """dialogue to change the venv used for an excel workbook"""
    new_python_interpreter = dlg_select_and_optionally_create_valid_python_interpreter(version_required=None)
    get_venv_root_directory_for_xlpro(new_python_interpreter)
    existing_xlpro_venv_root_path = get_valid_venv_root_path_used_for_workbook_from_map(workbook_path)
    print_info(f"removing existing binding to {existing_xlpro_venv_root_path} for {workbook_path}")
    remove_venv_to_workbook_mapping(workbook_path=workbook_path, venv_path=existing_xlpro_venv_root_path)
    store_venv_to_workbook_mapping(workbook_path=workbook_path, xlpro_venv_parent_path=existing_xlpro_venv_root_path)
    
    write_server_vscode_config_settings(workbook_path=workbook_path, active_venv=existing_xlpro_venv_root_path)


def get_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('localhost', 0))  # Binding to port 0 tells OS to assign a free port
        _, port = s.getsockname()
        return port


def check_port_old(host, port):
    try:
        # Create a socket object
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(5)  # Optional: Set timeout to avoid hanging forever
        # Try to connect to the given host and port
        sock.connect((host, port))
    except socket.error as e:
        return False  # Port is closed or unreachable
    finally:
        sock.close()
    return True  # Port is open

def check_port(host, port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        try:
            s.bind((host, port))
            return True
        except OSError:
            return False
        

def update_workbook_debugpy_port(workbook_path:Path, port:int) -> None:
    """Takes an xlpro venv interpreter and workbook path, 
    stores the interpreter path to debug port map, and updates the server launch.json for the correct port
    """
    xlpro_wd = get_xlpro_workbook_directory(workbook_path)
    print_info(f"updating debugpy port for {workbook_path} to {port}...")
    write_launch_json(xlpro_wd, port)
    print_success(f"updated debugpy port for {workbook_path} to {port} (.vscode/launch.json)")


def read_configurations_json(workbook_path:Path) -> int:
    d = read_json_file_with_comments(get_xlpro_workbook_directory(workbook_path) / ".vscode/launch.json")
    return d


def read_xlpro_debug_configuration_port(workbook_path:Path):
    d = read_configurations_json(workbook_path)
    xlpro_config = d["configurations"][0]
    if not xlpro_config["name"] == "xlpro debugpy":
        print_error(msg:=f"could not locate the xlpro configuration for {workbook_path}")
        raise Exception(msg)
    ret = xlpro_config["connect"]["port"]
    if not isinstance(ret, int):
        raise TypeError
    return ret


# XXX - todo unify this across xlpro_module...
def check_lockfile_get_contents_as_dict_if_alive(lock_file) -> dict:
    """Check if a process holding the lock is still running."""
    try:
        with open(lock_file, 'r') as f:
            contents = f.read().strip() 

            pid = re.search("pid=(.+)$", contents, re.MULTILINE).group(1)
            guid = re.search("guid=(.+)$", contents, re.MULTILINE).group(1)
            debugpy_port = re.search("debugpy_port=(.+)$", contents, re.MULTILINE).group(1)
            
            if psutil.pid_exists(pid):
                # Process is still running
                return {
                    "pid": pid, 
                    "guid": guid,
                    "debugpy_port":debugpy_port,
                }
    except (ValueError, FileNotFoundError):
        pass
    return {}

# XXX - todo unify this across xlpro_module...
def get_xlpro_lockfile_path(interpreter_path:Path=None) -> Path:
    if interpreter_path is not None:
        xlpro_dir = interpreter_path.parent.parent.parent / ".xlpro"
    else:
        xlpro_dir = Path(sys.executable).parent.parent.parent / ".xlpro"
    return xlpro_dir / "xlpro.lock"


def get_running_pid_guid_port_for_workbook(workbook_path:Path) -> tuple[int, str, int]|None:
    # get the interpreter from the workbook to determine the lockfile name
    xlpro_venv_root_path = get_valid_venv_root_path_used_for_workbook_from_map(workbook_path)
    interpreter_path = get_python_exe_from_xlpro_root_venv_path(xlpro_venv_root_path)

    # the lockfile path will be set adjacent to the interpreter running the xlpro_module
    # however when running using the debug venv interpreter, there will be a mismatch between the interpreter within xlpro-cli.exe and
    # the one which generated the lockfile. 
    # manual run_server calls are therefore now non-functional, until xlpro-cli.exe and the xlpro create the lockfile at a consistent location.
    # without the lockfile, there is no way for xlpro-cli to return the guid

    xlpro_lock_fp = file_lock.get_xlpro_lockfile_path_parent(interpreter_path=interpreter_path) / f"{workbook_path.name}.xlpro.lock"
    # xlpro_lock_fp = xlpro.file_lock.get_xlpro_lockfile_path_parent() / f"{workbook_path.name}.xlpro.lock"
    
    lockfile_contents_dict = file_lock.check_lockfile_get_contents_as_dict_if_alive(xlpro_lock_fp)
    # xlpro_lock_fp = get_xlpro_lockfile_path(interpreter_path=interpreter_path)
    # lockfile_contents_dict = check_lockfile_get_contents_as_dict_if_alive(xlpro_lock_fp)

    # if we get data back, the process is alive so we should use this data.
    if lockfile_contents_dict:
        pid, guid, debugpy_port = [lockfile_contents_dict.get(x) for x in ("pid", "guid", "debugpy_port")]
        return pid, guid, debugpy_port
    
    return 


def get_interpreter_pid(interpreter_path:Path):
    raise NotImplementedError
    check = get_running_pid_guid_port_for_workbook(interpreter_path)
    if check is not None:
        pid, guid, port = check
        return pid
    
def get_interpreter_guid(workbook_path:Path):
    """gets the guid of the running interpreter"""
    check = get_running_pid_guid_port_for_workbook(workbook_path)
    if check is not None:
        pid, guid, port = check
        return guid
    
def get_interpreter_port(interpreter_path:Path):
    raise NotImplementedError
    check = get_running_pid_guid_port_for_workbook(interpreter_path)
    if check is not None:
        pid, guid, port = check
        return port
        

def start_venv_xlpro_server_for_workbook(workbook_path:Path, do_kill_running:bool=True, do_register_wb:bool=True):
    """spins up the xlpro server on a port specified in the launch.json debug configuration"""
    from win32com.client import Dispatch, GetActiveObject

    py_interpreter_root_dir = get_valid_venv_root_path_used_for_workbook_from_map(workbook_path)

    if py_interpreter_root_dir is None:
        print_error(msg:=f"Interpreter was not found for {workbook_path}, please initialize first.")
        raise Exception(msg)

    py_interpreter_path = get_python_exe_from_xlpro_root_venv_path(py_interpreter_root_dir)
    # look for the current launch json configuration
    port = read_xlpro_debug_configuration_port(workbook_path)
    workbook_xlpro_wd = get_xlpro_workbook_directory(workbook_path)

    if do_kill_running:
        check = get_running_pid_guid_port_for_workbook(workbook_path=workbook_path)
        if check is None:
            print_info("No running server, ok to continue")
        else:
            _pid, _guid, _port = check
            print_info(f"Running server potentially running at pid: {_pid}")

            try:
                print_info(f"Attempting to close pid: {_pid}")
                process = psutil.Process(_pid)
                process.terminate()  # Graceful
                process.wait(timeout=3)
                print_success("Process closed successfully.")
            except psutil.NoSuchProcess:
                print_info("Process does not exist, ok to continue")
            except psutil.TimeoutExpired:
                print_error("Request timed out, could not close the server gracefully. Forcing...")
                process.kill()  # Force kill if it didn't terminate in time
                print_warning("Force close complete")
        
        
    if not check_port("localhost", port):
        print_warning(f"currently specified port {port} in launch.json is not available, finding another")
        port = get_free_port()
        print_success(f"free port found, {port}")
        # write_launch_json(workbook_xlpro_wd, port)
    
    # guid = None
    # # overwrite the port and guid
    # check = get_running_pid_guid_port_for_interpreter(py_interpreter_path)
    # if check:
    #     pid, guid, port = check
    #     print_info(f"Existing server found at pid:{pid}, guid:{guid}, port:{port}")

    update_workbook_debugpy_port(workbook_path=workbook_path, port=port)
    print_info(f"spinning up xlpro server for {py_interpreter_path} with debugpy port {port}")
    # if a python process already exists based on the lockfile, this will close itself!

    CONFIG = try_except_press_enter_to_exit_wrapper(config.load)()

    process = subprocess.Popen(
        [
            CONFIG.XLPRO_SERVER_PATH,
            str(py_interpreter_path),
            "-X",
            "frozen_modules=on",
            "-m",
            "xlpro.run_server",
            f"--debugpy_port={str(port)}",
            f"--workbook_path={str(workbook_path)}"
        ],
        stdin=subprocess.DEVNULL,
        creationflags=subprocess.CREATE_NEW_CONSOLE,
        text=True,
        # stdout=subprocess.PIPE,
        stderr=subprocess.PIPE # if do_register_wb else None, # Pipe the stderr to read the triggers
    )

    def register_wb_on_signal(_process):
        """Registers the workbook once the server is ready."""
        exit_message = "XLPROSTART_TRIGGER_OK"
        timeout_seconds = 10.0
        start_time = time.time()

        try:
            for line in _process.stderr:
                sys.stderr.write(line)  # Mirror stderr
                sys.stderr.flush()
                if exit_message in line:
                    break  # Trigger detected
                if time.time() - start_time > timeout_seconds:
                    raise TimeoutError("Timeout waiting for startup signal")
        except TimeoutError as e:
            print_error(f"Startup timed-out after {timeout_seconds} sec. You will need to manually register (sync) the workbook.")

        # Ready to link the workbook to the server.
        # Dispatch the workbook to run the registration macro from here 
        try:
            print_info("Signalling workbook to sync...")
            # xlapp = Dispatch("Excel.Application")
            xlapp = GetActiveObject("Excel.Application")
            wb = xlapp.Workbooks.Open(str(workbook_path))
            xlapp.Run("xlpro.xlam!xlproRegisterWorkbook", wb)
            print_info("Sync attempt complete.")
        except Exception as e:
            print_error("Could not signal to Excel to register the workbook. You will need to manually register (sync) the workbook.")

    if do_register_wb:
        register_wb_on_signal(process)

    press_enter_or_timeout_exit(1.0)


def get_folder_size(path: str | Path) -> int:
    path = Path(path)
    return sum(f.stat().st_size for f in path.rglob('*') if f.is_file())

def get_folder_size_fast(path):
    if isinstance(path, Path):
        path = str(path)
    total_size = 0
    with os.scandir(path) as it:
        for entry in it:
            try:
                if entry.is_file(follow_symlinks=False):
                    total_size += entry.stat().st_size
                elif entry.is_dir(follow_symlinks=False):
                    total_size += get_folder_size(entry.path)
            except (OSError, PermissionError):
                continue  # Skip inaccessible files
    return total_size

def delete_folder_onerror(func, path, exc_info):
    # Handle readonly or locked files
    os.chmod(path, stat.S_IWRITE)
    try:
        func(path)
    except Exception as e:
        print(f"Failed to delete {path}: {e}")

def delete_folder(path):
    if os.path.exists(path):
        shutil.rmtree(path, onerror=delete_folder_onerror)
        print(f"Deleted folder: {path}")
    else:
        print(f"Folder does not exist: {path}")

def check_folder_size_prompt_delete(fp:Path):
    print_info(f"Calculating folder size at {fp}...")
    if not fp.exists():
        print_info(f"Path not found at {fp}. No actions required")
        return
    size_bytes = get_folder_size(fp)
    # size_bytes = get_folder_size_fast(XLPRO_TMP_FOLDER_PATH)
    size_mb = size_bytes / (1024 * 1024)
    print_info(f"Folder size is {size_mb:.2f} MB")
    if prompt_yes_no_input("Do you want to remove this folder") == "yes":
        print_info("Deleting...")
        try:
            delete_folder(fp)
            new_size_bytes = get_folder_size(fp)
            new_size_mb = new_size_bytes / (1024 * 1024)
            print_success(f"Deleting finished, {size_mb} MB to {new_size_mb} MB")
        except Exception as e:
            print_error(f"Unable to delete the directory. {e}")



def add_to_user_path(p:Path):
    # Ensure the path is absolute
    new_path = p.absolute()
    str_new_path = str(new_path)
    try:
        # Open the registry key where user environment variables are stored
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment", 0, winreg.KEY_READ | winreg.KEY_WRITE) as key:
            
            # Get the current PATH value
            try:
                current_path, _ = winreg.QueryValueEx(key, "Path")
            except FileNotFoundError:
                current_path = ""

            # Check if the path is already in PATH
            if str_new_path in current_path.split(";"):
                print_info(f"The path '{new_path}' is already in the user PATH.")
                return

            # Append the new path
            updated_path = f"{current_path};{new_path}" if current_path else str_new_path

            # Write back to the registry
            winreg.SetValueEx(key, "Path", 0, winreg.REG_EXPAND_SZ, updated_path)
            print_success(f"Successfully added '{new_path}' to the user PATH.")
            
    except Exception as e:
        print_error(f"Error: {e}")


def remove_from_user_path(p:Path):
    # Ensure the path is absolute
    new_path = p.absolute()
    str_new_path = str(new_path)
    try:
        # Open the registry key where user environment variables are stored
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment", 0, winreg.KEY_READ | winreg.KEY_WRITE) as key:
            
            # Get the current PATH value
            try:
                current_path, _ = winreg.QueryValueEx(key, "Path")
            except FileNotFoundError:
                current_path = ""

            # Check if the path is already in PATH
            if not str_new_path in current_path.split(";"):
                print_info(f"The path '{new_path}' is already NOT in the user PATH.")
                return

            current_path_less_requested = current_path.split(";")
            current_path_less_requested.remove(str_new_path)

            # Append the new path
            updated_path = ";".join(current_path_less_requested)

            # Write back to the registry
            winreg.SetValueEx(key, "Path", 0, winreg.REG_EXPAND_SZ, updated_path)
            print_success(f"Successfully removed '{new_path}' from the user PATH.")
            
    except Exception as e:
        print_error(f"Error: {e}")





if __name__ == "__main__":
    # get_user_selection("select an option", ["a", "b", "c", "d"])
    # main()
    # main_but_reinitializing()
    # test_configure_workbook_with_existing_venv()

    from win32com.client.dynamic import Dispatch

    xlapp:"xl._Application" = Dispatch("Excel.Application")
    xlapp.Visible=False

    workbook_paths:list[Path] = []
    # workbook_paths.append(Path() / "xlpro_testing" / "test2-initialize-raw" / "book1.xlsx")
    # workbook_paths.append(Path() / "xlpro_testing" / "test2-initialize-from-existing-xlpro" / "book1.xlsx")
    # workbook_paths.append(Path() / "xlpro_testing" / "test2-initialize-from-uv" / "book1.xlsx")
    # workbook_paths.append(Path() / "xlpro_testing" / "test2-initialize-from-uv-download-option" / "book1.xlsx")
    # workbook_paths.append(Path() / "xlpro_testing" / "test2-initialize-from-uv-download-specify" / "book1.xlsx")


    def make_workbook(p:Path):
        p.parent.mkdir(exist_ok=True, parents=True)
        wb = xlapp.Workbooks.Add()
        wb.SaveAs(str(p.resolve()))
        wb.Close()
    
    for wb_fp in workbook_paths:
        print(f"wb_fp is {wb_fp}")
        make_workbook(wb_fp)
        dlg_xlpro_initialize_workbook(wb_fp)
        pass

    pass
    workbook_paths2:list[Path] = []
    workbook_paths2.append(Path() / "xlpro_testing" / "test3-initialize-from-existing-system" / "book1.xlsx")

    for wb_fp in workbook_paths2:
        print(f"wb_fp is {wb_fp}")
        make_workbook(wb_fp)
        dlg_xlpro_initialize_workbook(wb_fp)
        pass

    

    pass