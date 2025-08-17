import xlpro

# This file is the 'entry point' of xlpro to define functions in Excel.
# all functions must be 'registered' in this file *only*.

@xlpro.register_sub
def write_hello_from_python(activews:'xlpro.xlWorksheet'):
    """Writes a string to cell A1 of the *active* worksheet."""
    activews.Range("A1").Value2 = "Hello from Python!"

@xlpro.register_sub
def create_new_sheet(activewb:'xlpro.xlWorkbook'):
    """Creates a new worksheet"""
    activewb.Sheets.Add2()


pass

import subprocess
import sys
import xlpro

@xlpro.register_sub()
def do_something():
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


