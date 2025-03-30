from pathlib import Path
import winreg
import logging
import subprocess
import shutil


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
                print(f"The path '{new_path}' is already in the user PATH.")
                return

            # Append the new path
            updated_path = f"{current_path};{new_path}" if current_path else str_new_path

            # Write back to the registry
            winreg.SetValueEx(key, "Path", 0, winreg.REG_EXPAND_SZ, updated_path)
            print(f"Successfully added '{new_path}' to the user PATH.")
            
    except Exception as e:
        print(f"Error: {e}")

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
                print(f"The path '{new_path}' is already NOT in the user PATH.")
                return

            current_path_less_requested = current_path.split(";")
            current_path_less_requested.remove(str_new_path)

            # Append the new path
            updated_path = ";".join(current_path_less_requested)

            # Write back to the registry
            winreg.SetValueEx(key, "Path", 0, winreg.REG_EXPAND_SZ, updated_path)
            print(f"Successfully removed '{new_path}' from the user PATH.")
            
    except Exception as e:
        print(f"Error: {e}")


XLPRO_INSTALL_DIR = (Path(r"C:\Users\Daniel Evans") / ".xlpro").resolve()
XLPRO_TEMP_DIR = XLPRO_INSTALL_DIR / "tmp"
XLPRO_BIN_DIR = XLPRO_INSTALL_DIR / "bin"
XLPRO_ENVS_DIR = XLPRO_INSTALL_DIR / "envs"

def get_preinstalled_uv_path():
    return Path() / r"C:\Users\Daniel Evans\.local\bin\uv.exe"

def install_uv(download=True):
    if download:
        XLPRO_TEMP_DIR.mkdir(exist_ok=True)
        result1 = subprocess.run(
            [
                "curl",
                "-L",
                "-o",
                str(uv_zip_archive:=((XLPRO_TEMP_DIR / "uv.zip"))),
                "https://github.com/astral-sh/uv/releases/download/0.6.10/uv-x86_64-pc-windows-msvc.zip",
                
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
        print(f"could not remove {XLPRO_TEMP_DIR}")
        pass

def install():

    # if XLPRO_INSTALL_DIR.exists():
        # raise Exception

    XLPRO_INSTALL_DIR.mkdir(exist_ok=True)
    XLPRO_BIN_DIR.mkdir()
    XLPRO_TEMP_DIR.mkdir()
    XLPRO_ENVS_DIR.mkdir()
    
    """
    xlpro-cli.exe
    envs
    config.toml
    venv-mappings.json
    """

    # write template files
    (XLPRO_INSTALL_DIR / "venv-mappings.json").write_text("", "utf-8")
    (XLPRO_INSTALL_DIR / "config.toml").write_text("", "utf-8")

    # get xlpro-cli.exe binary
    src_xlpro_cli = Path() / "xlpro_cli" / "dist/xlpro-cli.exe" 
    dst_xlpro_cli = XLPRO_INSTALL_DIR / src_xlpro_cli.name
    shutil.copy2(src_xlpro_cli, dst_xlpro_cli)

    # add xlpro to user path so they can call xlpro-cli
    add_to_user_path(XLPRO_INSTALL_DIR)

    # install uv.exe in the /bin directory
    install_uv(download=False)


    # copy config.toml
    config_path = Path("config.toml")
    shutil.copy2(config_path, XLPRO_INSTALL_DIR / config_path.name)


if __name__ == "__main__":
    install()