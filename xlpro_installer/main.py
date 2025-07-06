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
import winreg
import logging
import subprocess
import shutil
import sys
import logging
import time

DEVELOPMENT_INSTALL = False

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


XLPRO_INSTALL_DIR = (Path(os.environ.get("USERPROFILE")) / ".xlpro").resolve()

XLPRO_XLAM_LOCALPATH = XLPRO_INSTALL_DIR / "src/xlpro.xlam"

XLPRO_TEMP_DIR = XLPRO_INSTALL_DIR / "tmp"
XLPRO_BIN_DIR = XLPRO_INSTALL_DIR / "bin"

IS_FROZEN = getattr(sys, 'frozen', False)

BASE_PATH = Path(__file__).parent
if getattr(sys, 'frozen', False):
    # Running in a PyInstaller bundle
    BASE_PATH = Path(sys._MEIPASS)
    # BASE_PATH = Path(sys.executable).parent.parent
    # BASE_PATH = Path(r"C:\Users\Daniel Evans\projects\xlpro\xlpro_installer")
else:
    # Running in a normal Python interpreter
    BASE_PATH = Path(__file__).parent

PRE_BUILD_DIR = BASE_PATH / "install"

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
        raise e


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

    new_uv_path.parent.mkdir(exist_ok=True, parents=True)
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

    errors = []

    logger.info("Installing xlpro")
    logger.info("Configuring installation directory")

    if XLPRO_INSTALL_DIR.exists():
        logger.critical(f"{XLPRO_INSTALL_DIR} already exists, please remove this folder if you wish to install.")
        raise FileExistsError()

    logger.info("(0/4) Copying Folder Structure...")
    shutil.copytree(PRE_BUILD_DIR, XLPRO_INSTALL_DIR)

    add_to_user_path(XLPRO_INSTALL_DIR)

    logger.info("Downloading uv...")
    # install uv.exe in the /bin directory
    if DEVELOPMENT_INSTALL:
        logger.info("Fetching local uv (DEVELOPMENT BUILD)")
        install_uv(download=False)
    else:
        logger.info("Downloading uv...")
        install_uv(download=True)


    try:
        logger.info("Copying xlpro.xlam file to XLSTART...")
        xlstart_path = get_xlstart_path()
        dst_xlpro_xlam_path2 = xlstart_path / XLPRO_XLAM_LOCALPATH.name
        if dst_xlpro_xlam_path2.exists():
            logger.warning(f"Warning {dst_xlpro_xlam_path2} already exists, not copying to xlstart")
            if prompt_yes_no_input("Do you want to delete the existing xlpro.xlam file?") == "yes":
                os.remove(dst_xlpro_xlam_path2)
                logger.info("Copying xlpro.xlam to XLSTART")
                shutil.copy2(XLPRO_XLAM_LOCALPATH, dst_xlpro_xlam_path2)
                logger.info("xlpro.xlam added successfully to XLSTART")
        else:
            logger.info("Copying xlpro.xlam to XLSTART")
            shutil.copy2(XLPRO_XLAM_LOCALPATH, dst_xlpro_xlam_path2)
            logger.info("xlpro.xlam added successfully to XLSTART")

    except FileNotFoundError as e:
        logger.critical(f"Could not locate XLSTART directory")
        msg = f"Could not locate XLSTART directory, {str(e)}"
        e = Exception(msg)
        logger.critical(e)
        errors.append(e)



    logger.info("Installation completed successfully.")

    return errors


from rich.console import Console
from rich.style import Style

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


def uninstall():
    errors = []
    logger.warning(f"xlpro appears to be installed at {XLPRO_INSTALL_DIR}")
    logger.warning(f"Prompting user to uninstall...")
    check = prompt_yes_no_input("Would you like to uninstall xlpro?", default="no")
    if check == "no":
        logger.info("Uninstall aborted.")
        return errors
    logger.info("(0/3) Checking if xlpro folder can be deleted...")
    success_check = can_delete_all(str(XLPRO_INSTALL_DIR))
    if not success_check:
        msg = f"Cannot delete xlpro folder at {XLPRO_INSTALL_DIR}, is it still being used?"
        logger.error(msg)
        e = Exception(msg)
        errors.append(e)
        logger.critical("Aborting uninstall")
        return errors
    try:
        logger.info(f"(1/3) Removing folder at {XLPRO_INSTALL_DIR}...")
        shutil.rmtree(XLPRO_INSTALL_DIR)
        logger.info("Folder removed")
    except Exception as e:
        msg = f"Error encountered during uninstall: {str(e)}"
        logger.error(msg)
        e = Exception(msg)
        errors.append(e)
        logger.critical("Aborting uninstall")
        return errors

    logger.info("(2/3) Attempting to install xlpro.xlam to XLSTART...")
    xlstart_path = get_xlstart_path()
    dst_xlpro_xlam_path2 = xlstart_path / "xlpro.xlam"
    if dst_xlpro_xlam_path2.exists():
        try:
            os.remove(dst_xlpro_xlam_path2)
        except Exception as e:
            msg = f"Unable to delete xlpro.xlam at {dst_xlpro_xlam_path2}. {str(e)}"
            logger.error(msg)
            e = Exception(msg)
            errors.append(e)
        
    logger.info("(3/3) Clearing user path key...")
    try:
        remove_from_user_path(XLPRO_INSTALL_DIR)
    except Exception as e:
        logger.error(f"Unable to remove xlpro from user path at {XLPRO_INSTALL_DIR}")
        errors.append(e)

    # errors.append(Exception("NOTE, files downloaded by uv have not be removed. Please uninstall the uv cache yourself if desired."))

    return errors

def main():
    errors = None
    try:
        if XLPRO_INSTALL_DIR.exists():
            errors = uninstall()
        else: 
            errors = install()
    except Exception as e:
        logger.critical(f"Fatal error encountered during installation: '{e}'")
    finally:
        if errors:
            logger.warning("Warning: Errors encountered:")
            for e in errors:
                logger.error(str(e))

        input("Press enter to exit.")
        sys.exit()

if __name__ == "__main__":
    main()