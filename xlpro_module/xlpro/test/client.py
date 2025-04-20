import win32com.client
import datetime, os, threading

from xlpro.server import xlproServer

# # now works with the progid after registering it in the registry
# xlpro = win32com.client.Dispatch(xlproServerAsync._reg_progid_)
# xlpro2 = win32com.client.Dispatch(xlproServerAsync._reg_progid_)

# # xlpro and xlpro2 are the same instance :)

# xlpro2 = win32com.client.Dispatch(xlproServerAsync._reg_clsid_)
# print(xlpro.add_data(f"my_pid(1): {os.getpid()}"))
# print(xlpro2.add_data(f"my_pid(2): {os.getpid()}"))
# # https://web.archive.org/web/20140917092059/http://www.devshed.com/c/a/Python/Windows-Programming-in-Python-Creating-COM-Servers/


# xlapp = win32com.client.Dispatch("Excel.Application")

# xlpro = win32com.client.Dispatch(xlproServer._reg_clsid_)
# # xlpro = win32com.client.Dispatch(xlproServerAsync._reg_progid_)

# ret1 = xlpro.getpid()
# print(ret1)
# xlpro.__dev_shutdown()


import subprocess
from pathlib import Path
import time
import sys
from win32com.client import Dispatch

wb_path = Path() / r"C:\Users\Daniel Evans\projects\xlpro\xlpro_testing\main.xlsx"

# subprocess.run(
#     [
#         r"C:\Users\Daniel Evans\projects\xlpro\.venv\Scripts\xlpro-cli.exe",
#         "start",
#         str(wb_path),
#     ],
#     text=True,
# )
# pass
sys.argv = ["commands.py", "start", str(wb_path)]
from xlpro_cli import commands
commands.main()

time.sleep(2)
process = subprocess.Popen(
    [
        r"C:\Users\Daniel Evans\projects\xlpro\.venv\Scripts\xlpro-cli.exe",
        "guid",
        str(wb_path),
    ],
    text=True,
    stderr=subprocess.PIPE,
)

sys.argv = ["guid", str(wb_path)]
for line in process.stderr:
    print(line, end='', file=sys.stderr)  # Print stderr immediately
    clsid = line

xlproserver:xlproServer = Dispatch(clsid)

pass


xlapp = Dispatch("Excel.Application")
wb = xlapp.Workbooks(wb_path.name)
xlproserver.register_and_configure_wb_workspace(wb)

pass



pass