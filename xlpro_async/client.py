import win32com.client
import datetime, os, threading

from server import xlproServerAsync

xlpro = win32com.client.Dispatch(xlproServerAsync._reg_clsid_)
xlpro2 = win32com.client.Dispatch(xlproServerAsync._reg_clsid_)

print(xlpro.add_data(f"my_pid(1): {os.getpid()}"))
print(xlpro2.add_data(f"my_pid(2): {os.getpid()}"))

xlpro2 = win32com.client.Dispatch(xlproServerAsync._reg_clsid_)
print(xlpro.add_data(f"my_pid(1): {os.getpid()}"))
print(xlpro2.add_data(f"my_pid(2): {os.getpid()}"))
# https://web.archive.org/web/20140917092059/http://www.devshed.com/c/a/Python/Windows-Programming-in-Python-Creating-COM-Servers/

pass