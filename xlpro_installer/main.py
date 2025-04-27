from pathlib import Path
import winreg
import logging
import subprocess
import shutil
import os
import sys
import logging
import version

DEVELOPMENT_BUILD = False

logging.basicConfig(
    # filename= wd / 'log.log',   # The file where logs will be saved
    # filename=config.logging_path,   # The file where logs will be saved
    stream=sys.stdout,
    # level=logging.DEBUG,          # The log level (DEBUG, INFO, WARNING, etc.)
    level=logging.INFO,          # The log level (DEBUG, INFO, WARNING, etc.)
    # format='%(asctime)s - %(levelname)s - %(message)s',  # The format of log messages
    format='%(levelname)s - %(message)s',  # The format of log messages
    # datefmt='%Y-%m-%d %H:%M:%S'    # The format of the date in log messages
)
logger =  logging.getLogger(__name__)

BASE_PATH = Path(__file__).parent
if getattr(sys, 'frozen', False):
    # Running in a PyInstaller bundle
    BASE_PATH = Path(sys._MEIPASS)
    # BASE_PATH = Path(sys.executable).parent.parent
    # BASE_PATH = Path(r"C:\Users\Daniel Evans\projects\xlpro\xlpro_installer")
else:
    # Running in a normal Python interpreter
    BASE_PATH = Path(__file__).parent


IS_FROZEN = getattr(sys, 'frozen', False)


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
                logger.debug(f"Requested path '{new_path}' is already in the user PATH.")
                return

            # Append the new path
            updated_path = f"{current_path};{new_path}" if current_path else str_new_path

            # Write back to the registry
            winreg.SetValueEx(key, "Path", 0, winreg.REG_EXPAND_SZ, updated_path)
            logger.info(f"Successfully added '{new_path}' to the user PATH.")
            
    except Exception as e:
        logger.warning(f"Failed to add to user path. '{e}'")

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
                logger.debug(f"The path '{new_path}' is already NOT in the user PATH.")
                return

            current_path_less_requested = current_path.split(";")
            current_path_less_requested.remove(str_new_path)

            # Append the new path
            updated_path = ";".join(current_path_less_requested)

            # Write back to the registry
            winreg.SetValueEx(key, "Path", 0, winreg.REG_EXPAND_SZ, updated_path)
            logger.debug(f"Successfully removed '{new_path}' from the user PATH.")
            
    except Exception as e:
        logger.debug(f"Error removing key: {e}")

# XLPRO_INSTALL_DIR = (Path(r"C:\Users\Daniel Evans") / ".xlpro").resolve()
XLPRO_INSTALL_DIR = (Path(os.environ.get("USERPROFILE")) / ".xlpro").resolve()

XLPRO_TEMP_DIR = XLPRO_INSTALL_DIR / "tmp"
XLPRO_BIN_DIR = XLPRO_INSTALL_DIR / "bin"
XLPRO_ENVS_DIR = XLPRO_INSTALL_DIR / "envs"
XLPRO_ASSETS_DIR = XLPRO_INSTALL_DIR / "assets"

XLPRO_INSTALLER_ASSETS_DIR = BASE_PATH / "assets"

def get_preinstalled_uv_path():
    return Path() / r"C:\Users\Daniel Evans\.local\bin\uv.exe"

def install_uv(download=True):
    if download:
        XLPRO_TEMP_DIR.mkdir(exist_ok=True)
        url = "https://github.com/astral-sh/uv/releases/download/0.6.10/uv-x86_64-pc-windows-msvc.zip"
        logging.info(f"Downloading uv from {url}")

        result1 = subprocess.run(
            [
                "curl",
                "-L",
                "-o",
                str(uv_zip_archive:=((XLPRO_TEMP_DIR / "uv.zip"))),
                url,
                
            ],
            check=True,
            capture_output=True
        )

        
        pass
        uv_extracted_dir = XLPRO_TEMP_DIR / "uv"
        uv_extracted_dir.mkdir()

        result2 = subprocess.run(
            [
                "tar",
                "-xf",
                str(uv_zip_archive),
                "-C",
                uv_extracted_dir
            ],
            check=True,
            capture_output=True
        )
        uv_exe_path = uv_extracted_dir / "uv.exe"

    else:
        logger.warning("Fetching local copy of uv instead of downloading")
        uv_exe_path = get_preinstalled_uv_path()

    new_uv_path = XLPRO_BIN_DIR / uv_exe_path.name
    
    if new_uv_path.exists():
        new_uv_path.unlink()

    shutil.copy2(uv_exe_path, new_uv_path)

    if download:
        uv_zip_archive.unlink()
        shutil.rmtree(uv_extracted_dir)

    try:
        XLPRO_TEMP_DIR.rmdir()
    except:
        logger.warning(f"Could not remove temporary directory {XLPRO_TEMP_DIR}")
        pass

def get_xlstart_path():
    user_profile = Path(os.environ.get('USERPROFILE'))
    if not user_profile.exists():
        raise EnvironmentError("USERPROFILE environment variable not found or invalid.")
    
    xlstart_path = user_profile / 'AppData/Roaming/Microsoft/Excel/XLSTART'
    
    if xlstart_path.is_dir():
        return xlstart_path
    else:
        raise FileNotFoundError(f"XLSTART folder not found at {xlstart_path}")


def install():

    # if XLPRO_INSTALL_DIR.exists():
        # raise Exception
    
    logger.info("Installing xlpro")
    logger.info("Configuring installation directory")
    if XLPRO_INSTALL_DIR.exists():
        logger.critical(f"{XLPRO_INSTALL_DIR} already exists, please remove this folder if you wish to install.")
        raise FileExistsError()

    logger.info("Creating subdirectories")
    XLPRO_INSTALL_DIR.mkdir()
    XLPRO_BIN_DIR.mkdir()
    XLPRO_TEMP_DIR.mkdir()
    XLPRO_ENVS_DIR.mkdir()
    XLPRO_ASSETS_DIR.mkdir()
    
    logger.info("Creating file templates")
    (XLPRO_ENVS_DIR / "venv-mappings.json").write_text("", "utf-8")
    (XLPRO_INSTALL_DIR / "config.toml").write_text("", "utf-8")

    if IS_FROZEN:
        src_xlpro_cli = BASE_PATH / "assets/xlpro-cli.exe"
        src_xlpro_xlam_path = BASE_PATH / "assets/xlpro.xlam"
        src_config_path = BASE_PATH / "assets/config.toml"
    else:
        src_xlpro_cli = BASE_PATH / "assets/xlpro-cli.exe"
        src_xlpro_xlam_path = BASE_PATH / "assets/xlpro.xlam"
        src_config_path = BASE_PATH / "assets/config.toml"

    logger.info("Fetching xlpro-cli binary")
    dst_xlpro_cli = XLPRO_INSTALL_DIR / src_xlpro_cli.name
    shutil.copy2(src_xlpro_cli, dst_xlpro_cli)

    logger.info("Adding to user path")
    # add xlpro to user path so they can call xlpro-cli
    add_to_user_path(XLPRO_INSTALL_DIR)

    logger.info("Downloading uv")
    # install uv.exe in the /bin directory
    if DEVELOPMENT_BUILD:
        logger.info("Fetching local uv (developer build)")
        install_uv(download=False)
    else:
        logger.info("Downloading uv (production build)")
        install_uv(download=True)

    logger.info("copying xlam file")
    dst_xlpro_xlam_path1 = XLPRO_ASSETS_DIR / src_xlpro_xlam_path.name
    shutil.copy2(src_xlpro_xlam_path, dst_xlpro_xlam_path1)

    try:
        logger.info("Attempting to install xlpro.xlam to XLSTART")
        xlstart_path = get_xlstart_path()
        dst_xlpro_xlam_path2 = xlstart_path / src_xlpro_xlam_path.name
        if dst_xlpro_xlam_path2.exists():
            logger.warning(f"warning {dst_xlpro_xlam_path2} already exists, not copying to xlstart")
        else:
            logger.info("Installing xlpro.xlam to XLSTART")
            shutil.copy2(dst_xlpro_xlam_path1, dst_xlpro_xlam_path2)
            logger.info("xlpro.xlam added successfully to XLSTART")

    except FileNotFoundError:
        logger.info(f"Could not locate XLSTART directory")


    # copy config.toml
    logger.info("Copying config")
    shutil.copy2(src_config_path, XLPRO_INSTALL_DIR / src_config_path.name)

    logger.info("Copying startfiles")
    shutil.copytree(x:=(XLPRO_INSTALLER_ASSETS_DIR / "startfiles"), XLPRO_ASSETS_DIR / x.name)    

    logger.info("Copying examples")
    shutil.copytree(x:=(XLPRO_INSTALLER_ASSETS_DIR / "examples"), XLPRO_ASSETS_DIR / x.name)       

    logger.info("Copying wheel")
    shutil.copy2(x:=(list(XLPRO_INSTALLER_ASSETS_DIR.glob("*.whl"))[0]), XLPRO_ASSETS_DIR / x.name)       

    # install xlpro.xlam
    pass
    logger.info("Installation completed successfully.")


from rich.console import Console
from rich.style import Style

console = Console(highlight=False)

style_prompt = Style.parse("green")
style_prompt_boldface = style_prompt + Style.parse("bold")

style_generic_option = Style.parse("cyan")
style_selected_option = style_generic_option + Style.parse("bold") + Style.parse("reverse")

style_plain = Style.parse("")
style_plain_boldface = style_plain + Style.parse("bold")

style_success = Style.parse("green")
style_success_boldface = style_success + Style.parse("bold")

style_error = Style.parse("red")
style_error_boldface = style_error + Style.parse("bold")

style_warning = Style(color="#FFA500")
style_warning_boldface = style_warning + Style.parse("bold")

def prompt_yes_no_input(prompt:str, default:str = "yes") -> str:
    if not default in ["yes", "no"]:
        raise Exception
    
    lookup = {
        "yes": "yes",
        "y": "yes",
        "no": "no",
        "n": "no",
        "": default
    }
    def print_prompt():
        console.print(f"{prompt} ", style=style_prompt_boldface, end="")
        console.print(f"[{'Y' if default=='yes' else 'y'}/{'N' if default=='no' else 'n'}]:", style=style_prompt)
        sys.stdout.flush()

    print_prompt()
    inp = input()
    while not (v:=inp.lower()) in lookup.keys():
        console.print(f"{v} not recognized", style=style_error)
        print_prompt()
        inp = input()
    
    ret = lookup[inp]
    if ret == "":
        console.print(default, style=style_plain)
    return lookup[inp]

def can_delete_all(path):
    all_ok = True
    files_to_delete = []

    for root, dirs, files in os.walk(path, topdown=False):
        for name in files:
            file_path = os.path.join(root, name)
            if os.access(file_path, os.W_OK):
                files_to_delete.append(file_path)
            else:
                logger.warning(f"Cannot delete (no write access): {file_path}")
                all_ok = False

        for name in dirs:
            dir_path = os.path.join(root, name)
            if not os.access(dir_path, os.W_OK | os.X_OK):
                logger.warning(f"Cannot delete directory (no write/execute access): {dir_path}")
                all_ok = False



    if all_ok:
        logger.info("All files and directories are deletable.")
        logger.info(f"{len(files_to_delete)} files will be deleted.")
        # logger.info("Files that would be deleted:")
        # for f in files_to_delete:
        #     logger.info(f"  {f}")
    else:
        logger.critical("Some files or directories cannot be deleted")

    return all_ok


def main():
    if XLPRO_INSTALL_DIR.exists():
        logger.warning(f"{XLPRO_INSTALL_DIR} already exists, promping user to uninstall")
        check = prompt_yes_no_input("Would you like to uninstall xlpro?", default="no")
        if check == "yes":
            success_check = can_delete_all(str(XLPRO_INSTALL_DIR))
            if not success_check:
                logger.error(f"Could not uninstall xlpro at {XLPRO_INSTALL_DIR}, is it still being used?")
                logger.critical("Uninstallation aborted")
            try:
                logger.info("Uninstallation starting")
                shutil.rmtree(XLPRO_INSTALL_DIR)
                # uninstall()
                logger.info("Uninstallation completed successfully")
            except Exception as e:
                logger.error(f"error encountered during uninstall: '{e}'")
                logger.critical("Uninstallation aborted")
            finally:
                input("Press enter to exit.")
                sys.exit()
        else:
            logger.info("Ending")
    try:
        install()
    except Exception as e:
        logger.critical(f"Fatal error encountered during installation: '{e}'")
        input("Press enter to exit.")
        sys.exit()

if __name__ == "__main__":
    main()