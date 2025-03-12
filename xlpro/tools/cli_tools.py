from rich.console import Console
from rich import print
from rich.style import Style
import readchar
import sys
import subprocess
import time
from pathlib import Path
import regex as re

def clear_line(count=1):
    for _ in range(count):
        sys.stdout.write("\033[A\r" + " " * 100 + "\r")  # Overwrite the line with spaces
        sys.stdout.flush()

pass


console = Console(highlight=False)

style_prompt = Style.parse("bold green")
style_generic_option = Style.parse("cyan")
style_selected_option = style_generic_option + Style.parse("bold") + Style.parse("reverse")

xlpro_path = Path() / "xlpro_install"
import uuid

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

selected_option = select_python_interpreter()
console.print(f"[bold green]You selected:[/] {selected_option}")
env_parent_dir_name = create_venv_path_name()

result = subprocess.run([selected_option, "--version"], shell=True, capture_output=True, text=True)

version_match = re.search(r"Python (\d+\.\d+\.\d+)", result.stdout.split("\n")[0])

chosen_py_version = version_match.group(1)

venv_path = xlpro_path / 'envs' / f"{env_parent_dir_name}_{chosen_py_version}" /'.venv'



result = subprocess.run(
    args=[
        "uv",
        "venv",
        f"{venv_path}",
    ],
    env={"PYTHON_EXE": selected_option},
    check=True,
)

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

pass