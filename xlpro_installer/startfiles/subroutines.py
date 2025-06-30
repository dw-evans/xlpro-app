import subprocess
import sys
import xlpro

xlpro.register_sub()
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


pass