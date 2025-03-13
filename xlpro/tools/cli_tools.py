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
        XLPRO_ROOT_PATH.mkdir(exist_ok=False)
    except Exception as e:
        raise e

    if not XLPRO_VENV_WORKBOOKS_MAP_JSON_FP.exists():
        with open(XLPRO_VENV_WORKBOOKS_MAP_JSON_FP, "w") as f:
            f.write(json.dumps({}))

    (XLPRO_ROOT_PATH / "envs").mkdir()
    

def create_venv_path_name():
    return str(uuid.uuid4())

def select_python_interpreter():
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
            return menu_items[index]  # Return selected item
        
        time.sleep(0.01)

def get_py_exe_version(py_interpreter_path:Path):
    str_py_interpreter_path = str(py_interpreter_path)
    result = subprocess.run([str_py_interpreter_path, "--version"], shell=True, capture_output=True, text=True)
    version_match = re.search(r"Python (\d+\.\d+\.\d+)", result.stdout.split("\n")[0])
    return version_match.group(1)

def create_venv_from_interpreter(py_interpreter_path:str):
    env_parent_dir_name = create_venv_path_name()
    chosen_py_version = get_py_exe_version(py_interpreter_path)
    venv_path = XLPRO_ENVS_DIR / f"{env_parent_dir_name}_{chosen_py_version}" /'.venv'
    result = subprocess.run(
        args=[
            "uv",
            "venv",
            f"{venv_path}",
        ],
        env={"PYTHON_EXE": py_interpreter_path},
        check=True,
    )
    initialize_xlpro_venv_files()

def initialize_xlpro_venv_files(venv_path:Path):
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

def initialize_workbook_for_xlpro(workbook_path:Path):
    from xlpro.server import configure_workspace_xlpro_files

    selected_py_interpreter = select_python_interpreter()
    configure_workspace_xlpro_files(workbook_path)
    store_venv_to_workbook_mapping(selected_py_interpreter)


def reinitialize_workbook_for_xlpro(workbook_path:Path) -> Path:
    """Reinitialize the environment for this workbook. Fetches the cached venv used"""
    found_venv = find_venv_used_for_workbook(workbook_path)

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

def write_requirements_txt_to_folder(environment_root_path:Path, parent_dir:Path):
    """Writes the requirements txt to an external folder"""
    result = subprocess.run(
        [
            "uv",
            "pip",
            "freeze",
            f"{str(parent_dir / 'requirements.txt')}",
        ],
        cwd=str(environment_root_path),
        capture_output=True, 
        shell=True,
        text=True,
    )


def read_json_file(fp:Path) -> dict:
    with open(fp, "w") as f:
        return json.loads(f.read())
    
def write_json_file(fp:Path, data:dict, indent=2):
    with open(fp, "w") as f:
        f.write(json.dumps(data, indent=indent))

    
def read_venv_to_workbooks_map() -> dict[str, list[str]]:
    return read_json_file(XLPRO_VENV_WORKBOOKS_MAP_JSON_FP)

def write_venv_to_workbooks_mappings_dict(venv_to_workbooks_map:dict) -> None:
    # venv mappings map the environment to the workbook, i.e. environment_path: workbook_path
    write_json_file(XLPRO_VENV_WORKBOOKS_MAP_JSON_FP, venv_to_workbooks_map)


def store_venv_to_workbook_mapping(workbook_path:str|Path, venv_path:str|Path):
    venv_to_workbooks_map = read_venv_to_workbooks_map()
    if not venv_path in venv_to_workbooks_map.keys():
        venv_to_workbooks_map[str(venv_path)] = []
    venv_to_workbooks_map[str(venv_path)].append(str(workbook_path))
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


def find_venv_used_for_workbook(workbook_path:str|Path) -> None|str|list[str]:
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
    
    return venv_to_workbooks_map_reversed[str_workbook_path]
    
def init_settings_json_python_path(parent_dir:Path, python_path:Path) -> Path:
    vscode_dir = (parent_dir / ".vscode")
    (parent_dir / ".vscode").mkdir(exist_ok=True)
    settings_dict = {}
    settings_dict["python.defaultInterpreterPath"] = python_path.resolve()
    settings_json_path = vscode_dir / "settings.json"
    write_json_file(settings_json_path, settings_dict)
    logger.warning(f"created settings.json at {settings_json_path}. Ensure you run the 'Python: Clear Workspace Interpreter' command from the command pallette")
    return settings_json_path

def create_python_version_file(interpreter_path:Path, parent_dir:Path):
    subprocess

if __name__ == "__main__":
    selected_option = select_python_interpreter()
    console.print(f"[bold green]You selected:[/] {selected_option}")
    create_venv_from_interpreter(selected_option)

pass