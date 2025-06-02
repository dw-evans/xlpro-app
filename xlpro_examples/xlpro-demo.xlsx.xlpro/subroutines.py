# > C:\Users\Daniel Evans\projects\xlpro\xlpro_testing\main.xlsx.xlpro\subroutines.py
# xlpro will automatically detect functions in this file as Excel subroutines.
import tkinter as tk
import subprocess
import sys
import xlpro
from pathlib import Path
import datetime

@xlpro.register_sub()
def do_something0():
    def spawn_console_window():
        """
        Spawns a non-blocking console window running a simple Python script
        that waits for the user to press Enter to exit.
        """
        python_code = 'input("Press Enter to exit...")'
        subprocess.Popen([
            sys.executable, "-c", python_code
        ], creationflags=subprocess.CREATE_NEW_CONSOLE)



    spawn_console_window()

@xlpro.register_sub()
def do_something():
    with open(Path(r"C:\Users\Daniel Evans\Downloads\MESSAGE\message.txt"), "a") as f:
        print(str(datetime.datetime.now()), file=f)

    pass


pass