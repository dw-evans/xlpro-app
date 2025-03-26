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


XLPRO_INSTALL_DIR = (Path() / "xlpro_install").resolve()
XLPRO_TEMP_DIR = XLPRO_INSTALL_DIR / "tmp"
XLPRO_BIN_DIR = XLPRO_INSTALL_DIR / "bin"



def install_uv():
    # result = subprocess.run(
    #     [
    #         "powershell.exe", 
    #         "-Command",
    #         "-ExecutionPolicy",
    #         "ByPass",
    #         "-c",
    #         "irm https://astral.sh/uv/0.6.10/install.ps1",
    #         "|",
    #         "iex",
    #     ]
    # )

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
        # check=True,
        capture_output=True
    )

    uv_exe_path = uv_extracted_dir / "uv.exe"

    new_uv_path = XLPRO_BIN_DIR / uv_exe_path.name
    if new_uv_path.exists():
        new_uv_path.unlink()

    shutil.copy2(uv_exe_path, new_uv_path)

    uv_zip_archive.unlink()
    shutil.rmtree(uv_extracted_dir)
    try:
        XLPRO_TEMP_DIR.rmdir()
    except:
        print(f"could not remove {XLPRO_TEMP_DIR}")
        pass


def install():
    pass

if __name__ == "__main__":
    install_uv()