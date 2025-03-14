from rich.console import Console
from rich import print
from rich.style import Style
import readchar
import sys
import subprocess
import time
from pathlib import Path
import regex as re
import logging
import json
import os


logging.basicConfig(
    stream=sys.stdout,
    level=logging.DEBUG,          # The log level (DEBUG, INFO, WARNING, etc.)
    format='%(asctime)s - %(levelname)s - %(message)s',  # The format of log messages
    datefmt='%Y-%m-%d %H:%M:%S'    # The format of the date in log messages
)

logger = logging.getLogger(__name__)



def clear_line(count=1):
    for _ in range(count):
        sys.stdout.write("\033[A\r" + " " * 100 + "\r")  # Overwrite the line with spaces
        sys.stdout.flush()

pass


console = Console(highlight=False)

style_prompt = Style.parse("bold green")
style_generic_option = Style.parse("cyan")
style_selected_option = style_generic_option + Style.parse("bold") + Style.parse("reverse")

XLPRO_ROOT_PATH = Path() / "xlpro_install"
XLPRO_VENV_WORKBOOKS_MAP_JSON_FP = XLPRO_ROOT_PATH / "venv-mappings.json"
XLPRO_ENVS_DIR = XLPRO_ROOT_PATH / 'envs'

import uuid

def initialize_xlpro_install_directory():
    try:
        XLPRO_ROOT_PATH.mkdir(exist_ok=True)
    except Exception as e:
        raise e

    if not XLPRO_VENV_WORKBOOKS_MAP_JSON_FP.exists():
        with open(XLPRO_VENV_WORKBOOKS_MAP_JSON_FP, "w") as f:
            f.write(json.dumps({}))

    (XLPRO_ROOT_PATH / "envs").mkdir(exist_ok=True)
    
def create_venv_path_name():
    return str(uuid.uuid4())

def select_python_interpreter() -> Path:
    py_path_locations = []
    
    result = subprocess.run("where.exe python", shell=True, capture_output=True, text=True)
    py_path_locations += result.stdout.split("\n")[:-1]

    result = subprocess.run("uv python dir", shell=True, capture_output=True, text=True)
    uv_py_dir = Path(result.stdout.split("\n")[0])
    uv_py_exes = [x for x in uv_py_dir.glob("*/python.exe")]
    py_path_locations += uv_py_exes

    menu_items = [f"{x}" for x in py_path_locations]
    menu_items += [
        # "(uv-download) Python 3.9", 
        # "(uv-download) Python 3.10", 
        # "(uv-download) Python 3.11", 
        # "(uv-download) Python 3.12", 
        # "(uv-download) Python 3.13",
        # "Other (specify)"
    ]
    index = 0
    s_list = []
    def print_options():
        nonlocal s_list
        s_list = []
        s = "Select your python interpreter:"
        console.print(s, style=style_prompt)
        s_list.append(s)
        for i, item in enumerate(menu_items):
            if i == index:
                s = f"> {item}  "
                console.print(s, style=style_selected_option)
            else:
                s = f"  {item}  "
                console.print(s, style=style_generic_option)
            s_list.append(s)
    
    print_options()

    while True:
        clear_line(len(s_list))
        print_options()

        key = readchar.readkey()

        if key == readchar.key.UP:
            index = (index - 1) % len(menu_items)  # Move up
        elif key == readchar.key.DOWN:
            index = (index + 1) % len(menu_items)  # Move down
        elif key == readchar.key.ENTER:
            return Path() / menu_items[index]  # Return selected item
        
        time.sleep(0.01)

def get_py_exe_version(py_interpreter_path:Path) -> str:
    str_py_interpreter_path = str(py_interpreter_path)
    result = subprocess.run([str_py_interpreter_path, "--version"], shell=True, capture_output=True, text=True)
    version_match = re.search(r"Python (\d+\.\d+\.\d+)", result.stdout.split("\n")[0])
    return version_match.group(1)

def create_xlpro_venv_from_interpreter_and_get_root_path(py_interpreter_path:Path) -> Path:
    """Creates a venv in the xlpro env folder and returns the path to the .venv parent directory"""
    str_py_interpreter_path = str(py_interpreter_path)
    env_parent_dir_name = create_venv_path_name()
    chosen_py_version = get_py_exe_version(py_interpreter_path)
    venv_path = XLPRO_ENVS_DIR / f"{env_parent_dir_name}_{chosen_py_version}" /'.venv'
    result = subprocess.run(
        args=[
            "uv",
            "venv",
            f"{venv_path}",
        ],
        env={"PYTHON_EXE": str(py_interpreter_path.resolve())},
        check=True,
    )
    initialize_xlpro_venv_files(venv_path)
    return venv_path


def create_symbolic_venv_from_existing_venv(existing_venv_root_path:Path) -> None:
    """Creates a symlinked virtual environment in the xlpro directory using a junction. Must provide the root
    path to the existing virtual environment"""

    if not (v:=(existing_venv_root_path / "scripts/python.exe")).exists():
        logger.error(f"python.exe does not exist within {existing_venv_root_path}")
        raise Exception(f"provided environment is invalid (no {v} found)")

    py_version = get_py_exe_version(xlpro_venv_path / "scripts/python.exe")
    env_parent_dir_name = create_venv_path_name()

    xlpro_venv_path = XLPRO_ENVS_DIR / f"{env_parent_dir_name}_{py_version}" /'.venv'

    result = subprocess.run(
        [
            "mkdir",
            "/J",
            ".venv",
            f"{existing_venv_root_path}",
            
        ],
        cwd=str(xlpro_venv_path.parent),
    )
    pass


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

    funcs_path = d / f"functions.py"
    subroutines_path = d / f"subroutines.py"

    p = funcs_path
    if not p.exists():
        with open(p, "w") as f:
            f.write(f"# > {p.resolve()}\n")
            f.write(f"# xlpro will automatically detect functions in this file as Excel UDFs.\n\n")

    p = subroutines_path
    if not p.exists():
        with open(p, "w") as f:
            f.write(f"# > {p.resolve()}\n")
            f.write(f"# xlpro will automatically detect functions in this file as Excel subroutines.\n\n")
    return d

def get_python_exe_from_root_venv_path(root_venv_path:Path) -> Path:
    return root_venv_path / "scripts/python.exe"


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


def scan_for_existing_xlpro_workbook_folder(workbook_path:Path) -> bool:
    potential_dir = get_xlpro_workbook_directory(workbook_path)
    if potential_dir.exists():
        validate_xlpro_workbook_folder(potential_dir)
        return True
    return False


def reinitialize_workbook_for_xlpro(workbook_path:Path) -> Path:
    """Reinitialize the environment for this workbook. Fetches the cached venv used"""
    found_venv = find_venv_root_path_used_for_workbook_from_map(workbook_path)

    # handle failed find 
    if not found_venv:
        logger.warning(f"the cached venv could not be found for {workbook_path}")
        remove_venv_to_workbook_mapping(workbook_path)
        raise Exception
    
    # handle venv does not exist
    if not Path(found_venv).exists():
        logger.warning(f"the cached venv '{found_venv}' does not exist")
        remove_venv_to_workbook_mapping()
        raise Exception

    return Path(found_venv)


def compare_environment_to_requirements_txt(environment_root_path:Path, external_requirements_txt_fp:Path) -> None:
    # environment_root_path = "."
    # external_requirements_txt_fp = "requirements.txt"
    result = subprocess.run(
        [
            "uv",
            "pip",
            "sync",
            "--dry-run"
            f"{str(external_requirements_txt_fp)}",
        ], 
        cwd=str(environment_root_path),
        capture_output=True, 
        shell=True,
        text=True,
    )
    lines = result.stderr.split("\n")[:-1]
    if lines[-1] == "Would make no changes":
        logger.debug(f"requirements.txt does match, ok to continue")
    else:
        logger.warning(f"requirements.txt does not match! \nmsg:\n{'  '.join(lines)}\n")
        input("press enter to update the environment per the above")
        print("Updating the environment...")
        result = subprocess.run(
            [
                "uv",
                "pip",
                "sync",
                f"{str(external_requirements_txt_fp)}",
            ], 
            cwd=str(environment_root_path),
            capture_output=True, 
            shell=True,
            text=True,
        )
        print("Updates complete")
    return

def compare_xlpro_venv_python_interpreter_version_to_required(venv_root:Path, workbook_path:Path):
    with open(get_xlpro_workbook_directory(workbook_path) / ".python-version", "r") as f:
        required_version = f.read() 
    active_version = get_py_exe_version(get_python_exe_from_root_venv_path(venv_root))
    logger.debug(f"required python version for {workbook_path} is {required_version}, active version is {active_version}")
    return active_version == required_version

def compare_venv_environment_to_required_environment(environment_root_path:Path, workbook_path:Path):
    compare_environment_to_requirements_txt(environment_root_path, get_xlpro_workbook_directory(workbook_path) / "requirements.txt")


def write_requirements_txt_to_folder(xlpro_venv_root_path:Path, xlpro_workbook_dir:Path) -> Path:
    """Writes the requirements txt to an external folder. Provide the parent folder relative to .venv
    for uv to handle dealing with the requirements.txt file.
    Returns the output requirements.txt file because why not"""
    outfile = xlpro_workbook_dir / 'requirements.txt'


    result = subprocess.run(
        [
            "uv",
            "pip",
            "freeze",
            "--python",
            str((xlpro_venv_root_path / ".venv/scripts/python.exe").resolve()),
            ">",
            str((xlpro_workbook_dir / 'requirements.txt').resolve()),
        ],
        cwd=str(xlpro_venv_root_path.resolve()), # PATH EXISTS
        capture_output=True, 
        shell=True,
        check=True,
        # text=True,
    )
    return outfile

def write_requirements_txt_for_workbook(xlpro_venv_root_path:Path, workbook_path):
    write_requirements_txt_to_folder(
        xlpro_venv_root_path,
        get_xlpro_workbook_directory(workbook_path),
    )

def read_json_file(fp:Path) -> dict:
    with open(fp, "r", encoding="utf-8") as f:
        try:
            d = json.loads(f.read())
        except: 
            d = {}
    return d
    
def write_json_file(fp:Path, data:dict, indent=2):
    with open(fp, "w", encoding="utf-8") as f:
        f.write(json.dumps(data, indent=indent))
    return
    
def read_venv_to_workbooks_map() -> dict[str, list[str]]:
    return read_json_file(XLPRO_VENV_WORKBOOKS_MAP_JSON_FP)

def write_venv_to_workbooks_mappings_dict(venv_to_workbooks_map:dict) -> None:
    # venv mappings map the environment to the workbook, i.e. environment_path: workbook_path
    write_json_file(XLPRO_VENV_WORKBOOKS_MAP_JSON_FP, venv_to_workbooks_map)


def store_venv_to_workbook_mapping(workbook_path:str|Path, xlpro_venv_parent_path:str|Path):
    """venv to workbook mapping maps the venv_root_dir.parent (xlpro custom name folder) to the workbook filepath
    i.e. one level above the /.venv folder..."""
    venv_to_workbooks_map = read_venv_to_workbooks_map()

    str_workbook_path = str(workbook_path)
    str_xlpro_venv_parent_path = str(xlpro_venv_parent_path)

    # construct reversed dict to check if the workbook isn't linked to a venv
    venv_to_workbooks_map_reversed:dict[str, str] = {}
    for k, v_list in venv_to_workbooks_map.items():
        for v in v_list:
            if v in venv_to_workbooks_map_reversed.keys():
                raise Exception("duplicated value, should not be possible")
            venv_to_workbooks_map_reversed[str(v)] = str(k)

    # remove the existing venv link if it exists per the above dict
    if str_workbook_path in venv_to_workbooks_map_reversed.keys():
        logger.warning(f"workbook {workbook_path} was previously mapped to {venv_to_workbooks_map_reversed[str_workbook_path]}")
        venv_to_workbooks_map[venv_to_workbooks_map_reversed[str_workbook_path]].remove(str_workbook_path)

    # add the new venv to workbook link as a new list item
    if not xlpro_venv_parent_path in venv_to_workbooks_map.keys():
        venv_to_workbooks_map[str_xlpro_venv_parent_path] = []
    venv_to_workbooks_map[str_xlpro_venv_parent_path].append(str_workbook_path)

    # save out the dict
    write_venv_to_workbooks_mappings_dict(venv_to_workbooks_map)


def remove_venv_to_workbook_mapping(workbook_path:str|Path, venv_path:str|Path|None=None) -> None:
    # todo - XXX - test this code
    logger.critical("untested code")

    venv_to_workbooks_map = read_venv_to_workbooks_map()

    str_workbook_path = str(workbook_path)
    str_venv_path = str(venv_path)

    # if only the workbook is given, delete the workbook and its links. E.g. the user wants to unlink a workbook
    if workbook_path and (venv_path is None):
        del venv_to_workbooks_map[str_workbook_path]
        write_venv_to_workbooks_mappings_dict()
        return
    # if only the venv is provided, remove any occurrences of the venv. E.g. it was deleted manually
    elif venv_path and (workbook_path is None):
        # todo - XXX - test this code
        for k, v, in venv_to_workbooks_map.items():
            if venv_path in v:
                venv_to_workbooks_map[k].remove(venv_path)
        write_venv_to_workbooks_mappings_dict()
        return

    # handle specific venv and workbook link removal
    if not str_venv_path in venv_to_workbooks_map.keys():
        logger.error(f"venv_path {str_venv_path} not found in {XLPRO_VENV_WORKBOOKS_MAP_JSON_FP.name}")
        return
    if not str_workbook_path in venv_to_workbooks_map[str_venv_path]:
        logger.error(f"workbook_path {str_workbook_path} not found in {XLPRO_VENV_WORKBOOKS_MAP_JSON_FP.name} key {str_venv_path}")

    venv_to_workbooks_map[str_venv_path].remove(str(workbook_path))
    write_venv_to_workbooks_mappings_dict(venv_to_workbooks_map)


def find_venv_root_path_used_for_workbook_from_map(workbook_path:str|Path) -> None|Path:
    """Finds the venv last used for the specified workbook if it exists, else returns None"""
    str_workbook_path = str(workbook_path)
    venv_to_workbooks_map = read_venv_to_workbooks_map()
    # venv_to_workbooks_map_reversed is {workbook_path: environment_path}
    venv_to_workbooks_map_reversed = {v: k for k, v_list in venv_to_workbooks_map.items() for v in v_list}
    if not str_workbook_path in venv_to_workbooks_map_reversed.keys():
        logger.error(f"workbook_path {str_workbook_path} was not found within in {XLPRO_VENV_WORKBOOKS_MAP_JSON_FP.name}. No environment found")
        return 
        # look for a filename match
        # todo - XXX - untested code
        wb_name = (Path() / str_workbook_path).name
        venv_to_workbooks_map_reversed_fname_only:dict[str, list[Path]] = {}
        # venv_to_workbooks_map_reversed_fname_only is {workbook_name: environment_path}
        for k, v_list in venv_to_workbooks_map_reversed.items():
            k_wb_name = Path(k).name
            if not k in venv_to_workbooks_map_reversed_fname_only:
                venv_to_workbooks_map_reversed_fname_only[k_wb_name] = []
            for v in v_list:
                venv_to_workbooks_map_reversed_fname_only[k_wb_name].append(Path(v))

        if k_wb_name in venv_to_workbooks_map_reversed_fname_only.keys():
            if val:=venv_to_workbooks_map_reversed_fname_only[k_wb_name]:
                logger.info(f"could not match the complete path {str_workbook_path} to a venv, but found {k_wb_name} was linked to {len(val)} items.")
                return val
        return 
    
    return Path() / venv_to_workbooks_map_reversed[str_workbook_path]


def write_settings_json_python_path(parent_dir:Path, python_exe_path:Path) -> Path:
    vscode_dir = (parent_dir / ".vscode")
    (parent_dir / ".vscode").mkdir(exist_ok=True)
    settings_dict = {}
    settings_dict["python.defaultInterpreterPath"] = str(python_exe_path.resolve())
    settings_json_path = vscode_dir / "settings.json"
    write_json_file(settings_json_path, settings_dict)
    logger.warning(f"created settings.json at {settings_json_path}. Ensure you run the 'Python: Clear Workspace Interpreter' command from the command pallette")
    return settings_json_path

def write_python_version_file(interpreter_path:Path, parent_dir:Path):
    with open(parent_dir / ".python-version", "w") as f:
        f.write(get_py_exe_version(interpreter_path))

def standardize_venv_root_to_parent(venv_root_path:Path) -> Path:
    if (venv_root_path / ".venv/scripts/python.exe").exists():
        return venv_root_path
    elif (venv_root_path / "scripts/python.exe").exists():
        return venv_root_path.parent
    raise Exception("invalid venv root path provided")



def main():
    """
    Layout:
        A user has a chosen workbook and calls xlpro init workbook_path 
        Steps:
            Prompt the user to 
                create a virtual environment
                (select a virtual environment and create a symlinked venv)
            init the local xlpro virtual environment
            init the adjacent xlpro directory
            create the settings.json
            ...
            on any change:
                write the requirements.txt to the xlpro destination directory
    """
    initialize_xlpro_install_directory()
    workbook_path = Path() / "xlpro_testing/test1/Book1.xlsx"
    workbook_path.parent.mkdir(exist_ok=True)
    workbook_path.write_text("", encoding="utf-8")

    # prompt the user to select the interpreter
    selected_py_interpreter = select_python_interpreter()
    # create the xlpro venv
    venv_root_path = create_xlpro_venv_from_interpreter_and_get_root_path(selected_py_interpreter)
    venv_exe_path = get_python_exe_from_root_venv_path(venv_root_path)

    # initialize the xlpro directory adjacent to the workbook
    xlpro_dir = initialize_and_get_workspace_xlpro_dir(workbook_path)
    # write the settings.json to the xlpro directory
    write_settings_json_python_path(xlpro_dir, venv_exe_path)

    subprocess.run(["where", "uv"], capture_output=True)
    subprocess.run(["uv", "pip", "list"], capture_output=True)

    # env = os.environ.copy()
    # env["PYTHON_EXE"] = str(venv_exe_path.resolve())

    subprocess.run(
        [
            "uv",
            "pip",
            "install",
            "--python",
            str(venv_exe_path.resolve()),
            "matplotlib",
            "pywin32",
        ],
        cwd=str(standardize_venv_root_to_parent(venv_root_path).resolve()),
        check=True,
        capture_output=True,
    )

    pass
    standardized_venv_root_dir = standardize_venv_root_to_parent(venv_root_path)

    # write_requirements_txt_for_workbook(xlpro_venv_root_path=standardize_venv_root_to_parent(venv_root_path), workbook_path=workbook_path)
    write_requirements_txt_to_folder(standardized_venv_root_dir, xlpro_dir)
    write_python_version_file(get_python_exe_from_root_venv_path(standardized_venv_root_dir / ".venv"), xlpro_dir)
    store_venv_to_workbook_mapping(workbook_path=workbook_path, xlpro_venv_parent_path=standardized_venv_root_dir)


    env = os.environ.copy()
    env["PYTHON_EXE"] = str(venv_exe_path.resolve())
    result = subprocess.run(
        [
            "uv",
            "pip",
            "uninstall",
            "matplotlib",
        ],
        cwd=str(standardize_venv_root_to_parent(venv_root_path).resolve()),
        env=env,
        check=True,
    )


    pass



def main_but_reinitializing():
    workbook_path = Path() / "xlpro_testing/test1/Book1.xlsx"
    workbook_path.parent.mkdir(exist_ok=True)
    workbook_path.write_text("", encoding="utf-8")

    is_xlpro = scan_for_existing_xlpro_workbook_folder(workbook_path=workbook_path)

    if not is_xlpro:
        raise Exception("oops make the file first dummy")

    # retrieve the venv from the cache
    venv_root_path = find_venv_root_path_used_for_workbook_from_map(workbook_path)
    if venv_root_path is None:
        raise Exception("No venv cound be found")

    # compare the requirements.txt
    compare_venv_environment_to_required_environment(venv_root_path)
    compare_xlpro_venv_python_interpreter_version_to_required(venv_root=venv_root_path / ".venv")

    store_venv_to_workbook_mapping(workbook_path=workbook_path, xlpro_venv_parent_path=venv_root_path.parent)

    pass





if __name__ == "__main__":
    main()
    # main_but_reinitializing()
    pass

pass