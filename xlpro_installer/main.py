from pathlib import Path
import winreg
import logging
import subprocess
import shutil
import os
import sys
import logging
import version

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
    # BASE_PATH = Path(sys._MEIPASS)
    # BASE_PATH = Path(sys.executable).parent.parent
    BASE_PATH = Path(r"C:\Users\Daniel Evans\projects\xlpro\xlpro_installer")
else:
    # Running in a normal Python interpreter
    BASE_PATH = Path(__file__).parent



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

XLPRO_INSTALL_DIR = (Path(r"C:\Users\Daniel Evans") / ".xlpro").resolve()
XLPRO_TEMP_DIR = XLPRO_INSTALL_DIR / "tmp"
XLPRO_BIN_DIR = XLPRO_INSTALL_DIR / "bin"
XLPRO_ENVS_DIR = XLPRO_INSTALL_DIR / "envs"
XLPRO_ASSETS_DIR = XLPRO_INSTALL_DIR / "assets"

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
    (XLPRO_INSTALL_DIR / "venv-mappings.json").write_text("", "utf-8")
    (XLPRO_INSTALL_DIR / "config.toml").write_text("", "utf-8")

    logger.info("Fetching xlpro-cli binary")
    # get xlpro-cli.exe binary
    src_xlpro_cli = BASE_PATH / "../xlpro_cli" / "dist/xlpro-cli.exe"
    dst_xlpro_cli = XLPRO_INSTALL_DIR / src_xlpro_cli.name
    shutil.copy2(src_xlpro_cli, dst_xlpro_cli)

    logger.info("Adding to user path")
    # add xlpro to user path so they can call xlpro-cli
    add_to_user_path(XLPRO_INSTALL_DIR)

    logger.info("Downloading uv")
    # install uv.exe in the /bin directory
    install_uv(download=False)
    # install_uv(download=True)

    logger.info("Attempting to install xlpro.xlam to XLSTART")
    # install xlpro.xlam

    
    src_xlpro_xlam_path = BASE_PATH / "../addin/xlpro.xlam"
    xlstart_path = get_xlstart_path()

    dst_xlpro_xlam_path1 = XLPRO_ASSETS_DIR / src_xlpro_xlam_path.name

    shutil.copy2(src_xlpro_xlam_path, dst_xlpro_xlam_path1)

    dst_xlpro_xlam_path2 = xlstart_path / src_xlpro_xlam_path.name
    if dst_xlpro_xlam_path2.exists():
        logger.warning(f"warning {dst_xlpro_xlam_path2} already exists, not copying to xlstart")
    else:
        logger.info("Installing xlpro.xlam to XLSTART")
        shutil.copy2(dst_xlpro_xlam_path1, dst_xlpro_xlam_path2)

    # copy config.toml
    logger.info("Copying config")
    config_path = BASE_PATH / "../config.toml"
    shutil.copy2(config_path, XLPRO_INSTALL_DIR / config_path.name)

    logger.info("Installation completed successfully.")

def main():
    try:
        install()
    except Exception as e:
        logger.critical(f"Fatal error encountered: '{e}'")
        input("Press enter to exit.")
        sys.exit(1)

if __name__ == "__main__":
    main()